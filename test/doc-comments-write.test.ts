// Comments in the docx writers: the comments part and the part of answers are
// made when a document first needs them, brought up to date when it has them,
// and left byte for byte where nothing was changed; undoing gives the file
// that was opened.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { buildDocx as convertToDocx, emptyWriteReport, writeDocx } from '../src/formats/docx/write';
import type { DocModel } from '../src/model/doc/types';
import { DocHistory, addComment, editComment, insertText, listComments, removeComment, replyToComment, resolveComment, setTracking, type DocRange } from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { parseXml } from '../src/xml';
import { EXTENDED_REL, brief, cEnd, cRef, cStart, checkCommentedPackage, commentXml, openCommented, sameCommentsAfterReading, type CommentedSpec } from './doc-comments-fixtures';
import { p, pos, r, range, run } from './doc-ops-fixtures';
import { buildDocx as buildSpec } from './docx-fixtures';
import { read as readRtf } from './rtf-fixtures';

const decoder = new TextDecoder();
const NOW = () => new Date('2026-10-12T08:09:10Z');
const COMMENTS_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
const EXTENDED_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';

async function open(spec: CommentedSpec) {
  const opened = await openCommented(spec);
  setTracking(opened.doc, { author: 'Me', now: NOW });
  const history = new DocHistory<DocRange | undefined>(opened.doc);
  const save = async (): Promise<Uint8Array> => {
    await writeDocx(opened.pkg, opened.doc);
    return opened.pkg.save();
  };
  const part = async (name: string): Promise<string> => decoder.decode(await opened.pkg.readPart(name));
  /** Every part of the package is what the file had, and there is no other part. */
  const sameAsOpened = async (): Promise<void> => {
    const pkg = await OpcPackage.open(await opened.pkg.save());
    assert.deepEqual(pkg.partNames().sort(), Array.from(opened.original.keys()).sort());
    for (const [name, bytes] of opened.original) assert.deepEqual(await pkg.readPart(name), bytes, name);
    assert.deepEqual(await opened.pkg.save(), opened.bytes, 'the saved file is not the file that was opened, byte for byte');
  };
  return { ...opened, history, save, part, sameAsOpened };
}

const ONE = p(r('One ') + cStart(1) + r('two') + cEnd(1) + cRef(1) + r(' three.')) + p('Second paragraph here.');
const THREAD: CommentedSpec = {
  body: p(r('One ') + cStart(1) + cStart(2) + r('two') + cEnd(1) + cRef(1) + cEnd(2) + cRef(2) + r(' three ') + cStart(3) + r('four') + cEnd(3) + cRef(3)) + p('Tail.'),
  comments: [
    { id: 1, author: 'Ann', text: 'First', paraId: '0000000A' },
    { id: 2, author: 'Bob', text: 'Answer', paraId: '0000000B' },
    { id: 3, author: 'Cid', text: 'Other', paraId: '0000000C' },
  ],
  extended: [['0000000A'], ['0000000B', '0000000A'], ['0000000C']],
};

test('the first comment of a document makes the comments part, its content type and its relationship', async () => {
  const o = await open({ body: p('One two three.') });
  assert.ok(!o.pkg.hasPart('/word/comments.xml'), 'the fixture has a comments part already');
  const edit = addComment(o.doc, range(pos(0, 4), pos(0, 7)), 'Why <two> & "three"?');
  run(o.history, edit);
  const report = emptyWriteReport();
  await writeDocx(o.pkg, o.doc, { report });
  const bytes = await o.pkg.save();
  assert.ok(report.parts.indexOf('/word/comments.xml') >= 0, 'the report does not name the new part');
  const pkg = await OpcPackage.open(bytes);
  assert.equal(pkg.contentTypeOf('/word/comments.xml'), COMMENTS_TYPE);
  const rels = (await pkg.relationships('/word/document.xml')).filter((rel) => /\/comments$/.test(rel.type));
  assert.deepEqual(rels.map((rel) => [rel.part, rel.type]), [['/word/comments.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments']]);
  const root = parseXml(decoder.decode(await pkg.readPart('/word/comments.xml'))).root;
  assert.equal(root.attrs.get('xmlns:w14'), W14);
  assert.ok((root.attrs.get('mc:Ignorable') || '').split(' ').indexOf('w14') >= 0, 'the root does not tell older readers to ignore w14');
  const facts = await checkCommentedPackage(bytes);
  assert.deepEqual(facts.ids, [edit.id]);
  assert.match(facts.paraIds.get(edit.id as number) as string, /^[0-9A-F]{8}$/);
  // Nothing is an answer and nothing is resolved: no part for that.
  assert.ok(!pkg.hasPart('/word/commentsExtended.xml'), 'a part of answers was made for a document without answers');
  const body = decoder.decode(await pkg.readPart('/word/document.xml'));
  assert.match(body, new RegExp(`<w:commentRangeStart w:id="${edit.id}"/>.*<w:commentRangeEnd w:id="${edit.id}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${edit.id}"/></w:r>`));
  const again = await sameCommentsAfterReading(bytes, o.doc);
  assert.equal(listComments(again)[0].text, 'Why <two> & "three"?');
  assert.deepEqual(again.warnings, []);
});

test('add a comment, undo, save: the file that was opened, byte for byte; also after a save in between', async () => {
  for (const spec of [{ body: p('One two three.') }, { body: ONE, comments: [{ id: 1, text: 'x' }] }, THREAD] as CommentedSpec[]) {
    const o = await open(spec);
    run(o.history, addComment(o.doc, range(pos(0, 0), pos(0, 3)), 'new'));
    o.history.undo();
    await o.save();
    await o.sameAsOpened();
    // The same with a save while the comment was there.
    o.history.redo();
    await checkCommentedPackage(await o.save());
    o.history.undo();
    await o.save();
    await o.sameAsOpened();
  }
});

test('answer, resolve, edit and remove, each undone after a save, leave the file that was opened', async () => {
  const steps: [string, (doc: DocModel) => ReturnType<typeof addComment>][] = [
    ['reply', (doc) => replyToComment(doc, 3, 'answer')],
    ['resolve', (doc) => resolveComment(doc, 1, true)],
    ['edit', (doc) => editComment(doc, 2, 'changed')],
    ['remove', (doc) => removeComment(doc, 1)],
  ];
  for (const [name, make] of steps) {
    const o = await open({ ...THREAD, ids: [['0000000A', '10000001'], ['0000000B', '10000002'], ['0000000C', '10000003']], extensible: [['10000001', '2026-10-01T09:00:00Z'], ['10000002', '2026-10-01T09:00:00Z']] });
    run(o.history, make(o.doc));
    const bytes = await o.save();
    await checkCommentedPackage(bytes);
    await sameCommentsAfterReading(bytes, o.doc);
    o.history.undo();
    await o.save();
    await o.sameAsOpened().catch((error) => {
      throw new Error(`${name}: ${error.message}`);
    });
  }
});

test('a comment that was not touched is written with the bytes it had; the others around it are added', async () => {
  const o = await open({ body: ONE, comments: [{ id: 1, author: 'Ann', text: 'Why two?', paraId: '0000000A', attrs: 'w:custom="kept"' }] });
  const original = decoder.decode(o.original.get('/word/comments.xml') as Uint8Array);
  const raw = commentXml({ id: 1, author: 'Ann', text: 'Why two?', paraId: '0000000A', attrs: 'w:custom="kept"' });
  assert.ok(original.indexOf(raw) >= 0, 'the fixture does not hold the comment as it was specified');
  const edit = addComment(o.doc, range(pos(1, 0), pos(1, 6)), 'Second');
  run(o.history, edit);
  const bytes = await o.save();
  const text = await o.part('/word/comments.xml');
  assert.ok(text.indexOf(raw) >= 0, 'the untouched comment was not copied from the source');
  assert.equal(text.slice(0, text.indexOf('<w:comment ')), original.slice(0, original.indexOf('<w:comment ')), 'the start of the part changed');
  const body = await o.part('/word/document.xml');
  const first = decoder.decode(o.original.get('/word/document.xml') as Uint8Array);
  assert.ok(body.indexOf(first.slice(first.indexOf('<w:p>'), first.indexOf('</w:p>') + 6)) >= 0, 'the paragraph that was not touched was not copied from the source');
  const facts = await checkCommentedPackage(bytes);
  assert.deepEqual(facts.ids, [1, edit.id]);
  await sameCommentsAfterReading(bytes, o.doc);
});

test('an answer makes the part of answers, with its content type and relationship, and gives the first comment a paragraph id', async () => {
  const o = await open({ body: ONE, comments: [{ id: 1, author: 'Ann', text: 'Why two?' }] });
  const reply = replyToComment(o.doc, 1, 'Because.');
  run(o.history, reply);
  const bytes = await o.save();
  const pkg = await OpcPackage.open(bytes);
  assert.equal(pkg.contentTypeOf('/word/commentsExtended.xml'), EXTENDED_TYPE);
  const rels = (await pkg.relationships('/word/document.xml')).filter((rel) => rel.type === EXTENDED_REL);
  assert.deepEqual(rels.map((rel) => rel.part), ['/word/commentsExtended.xml']);
  const facts = await checkCommentedPackage(bytes);
  const first = facts.paraIds.get(1) as string;
  const second = facts.paraIds.get(reply.id as number) as string;
  assert.match(first, /^[0-9A-F]{8}$/);
  assert.deepEqual(Array.from(facts.extended as Map<string, unknown>), [
    [first, { done: false }],
    [second, { parent: first, done: false }],
  ]);
  const root = parseXml(decoder.decode(await pkg.readPart('/word/commentsExtended.xml'))).root;
  assert.equal(root.attrs.get('mc:Ignorable'), 'w15');
  assert.equal(root.attrs.get('xmlns:w15'), 'http://schemas.microsoft.com/office/word/2012/wordml');
  const again = await sameCommentsAfterReading(bytes, o.doc);
  assert.deepEqual(brief(listComments(again)), [
    [1, 'Ann', 'Why two?', undefined, false],
    [reply.id as number, 'Me', 'Because.', 1, false],
  ]);
});

test('resolving changes one attribute of the part of answers and keeps everything else it holds', async () => {
  const o = await open(THREAD);
  // Something the engine does not know, and an entry for no comment at all.
  const extra = '<w15:future w15:x="1"/><w15:commentEx w15:paraId="0FFFFFFF" w15:done="0" w15:kept="yes"/>';
  const original = decoder.decode(o.original.get('/word/commentsExtended.xml') as Uint8Array).replace('</w15:commentsEx>', extra + '</w15:commentsEx>');
  const bytes = await buildSpec({ ...(await import('./doc-comments-fixtures')).commentedSpec(THREAD), parts: [['word/commentsExtended.xml', original]] });
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  const history = new DocHistory<DocRange | undefined>(doc);
  run(history, resolveComment(doc, 3, true));
  await writeDocx(pkg, doc);
  const text = decoder.decode(await pkg.readPart('/word/commentsExtended.xml'));
  assert.equal(text, original.replace('<w15:commentEx w15:paraId="0000000C" w15:done="0"/>', '<w15:commentEx w15:paraId="0000000C" w15:done="1"/>'));
  // The comments themselves say nothing new: their bytes are the file's.
  assert.equal(decoder.decode(await pkg.readPart('/word/comments.xml')).indexOf('w14:paraId="0000000C"') >= 0, true);
  history.undo();
  await writeDocx(pkg, doc);
  assert.deepEqual(await pkg.save(), bytes);
});

test('removing a comment takes its entries out of the parts beside the comments', async () => {
  const o = await open({ ...THREAD, ids: [['0000000A', '10000001'], ['0000000B', '10000002'], ['0000000C', '10000003']], extensible: [['10000001', '2026-10-01T09:00:00Z'], ['10000002', '2026-10-01T09:00:00Z'], ['10000003', '2026-10-01T09:00:00Z']] });
  run(o.history, removeComment(o.doc, 1));
  const bytes = await o.save();
  const facts = await checkCommentedPackage(bytes);
  assert.deepEqual(facts.ids, [3]);
  assert.deepEqual(Array.from((facts.extended as Map<string, unknown>).keys()), ['0000000C']);
  assert.doesNotMatch(await o.part('/word/commentsIds.xml'), /0000000A|0000000B/);
  assert.match(await o.part('/word/commentsIds.xml'), /w16cid:paraId="0000000C" w16cid:durableId="10000003"/);
  assert.doesNotMatch(await o.part('/word/commentsExtensible.xml'), /10000001|10000002/);
  assert.match(await o.part('/word/commentsExtensible.xml'), /10000003/);
  assert.doesNotMatch(await o.part('/word/document.xml'), /w:id="1"|w:id="2"/);
  await sameCommentsAfterReading(bytes, o.doc);
});

test('a new comment gets an entry in commentsIds.xml when the file has that part, and no such part is made', async () => {
  const withIds = await open({ ...THREAD, ids: [['0000000A', '10000001'], ['0000000B', '10000002'], ['0000000C', '10000003']] });
  const edit = addComment(withIds.doc, range(pos(1, 0), pos(1, 4)), 'new');
  run(withIds.history, edit);
  const facts = await checkCommentedPackage(await withIds.save());
  const paraId = facts.paraIds.get(edit.id as number) as string;
  const ids = await withIds.part('/word/commentsIds.xml');
  const match = new RegExp(`<w16cid:commentId w16cid:paraId="${paraId}" w16cid:durableId="([0-9A-F]{8})"/>`).exec(ids);
  assert.ok(match, 'the new comment has no entry in commentsIds.xml');
  assert.ok(parseInt((match as RegExpExecArray)[1], 16) < 0x7fffffff && ['10000001', '10000002', '10000003'].indexOf((match as RegExpExecArray)[1]) < 0, 'the durable id is not a new one');
  // The part of answers has it too: the file keeps one entry for each comment.
  assert.ok((facts.extended as Map<string, unknown>).has(paraId), 'the new comment has no entry in commentsExtended.xml');

  const without = await open(THREAD);
  run(without.history, addComment(without.doc, range(pos(1, 0), pos(1, 4)), 'new'));
  await without.save();
  assert.ok(!without.pkg.hasPart('/word/commentsIds.xml'), 'commentsIds.xml was made for a file that has none');
  assert.ok(!without.pkg.hasPart('/word/commentsExtensible.xml'), 'commentsExtensible.xml was made for a file that has none');
});

test('a comments part that does not declare the namespace of paragraph ids gets the declaration with the first id', async () => {
  const bare = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' + commentXml({ id: 1, text: 'Bare' }) + '</w:comments>';
  const bytes = await buildSpec({ body: ONE, parts: [['word/comments.xml', bare]], rels: [['rIdC', 'comments', 'comments.xml']] });
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  const history = new DocHistory<DocRange | undefined>(doc);
  run(history, resolveComment(doc, 1, true));
  await writeDocx(pkg, doc);
  const saved = await pkg.save();
  const root = parseXml(decoder.decode(await pkg.readPart('/word/comments.xml'))).root;
  const prefix = Array.from(root.attrs).find(([, value]) => value === W14);
  assert.ok(prefix, 'the namespace of paragraph ids is not declared');
  const name = (prefix as [string, string])[0].slice(6);
  const mc = Array.from(root.attrs).find(([, value]) => value === 'http://schemas.openxmlformats.org/markup-compatibility/2006');
  assert.ok(mc, 'the namespace of mc:Ignorable is not declared');
  assert.ok((root.attrs.get(`${(mc as [string, string])[0].slice(6)}:Ignorable`) || '').split(' ').indexOf(name) >= 0, 'the prefix is not among the ones to ignore');
  const facts = await checkCommentedPackage(saved);
  assert.deepEqual(Array.from(facts.extended as Map<string, { done: boolean }>).map(([, entry]) => entry.done), [true]);
  assert.equal(listComments(await readDocx(await OpcPackage.open(saved)))[0].done, true);
  history.undo();
  await writeDocx(pkg, doc);
  assert.deepEqual(await pkg.save(), bytes);
});

test('a document whose comments were not touched is not written at all, whatever else was edited', async () => {
  const o = await open(THREAD);
  await writeDocx(o.pkg, o.doc);
  await o.sameAsOpened();
  run(o.history, insertText(o.doc, pos(1, 0), 'Typed '));
  const report = emptyWriteReport();
  await writeDocx(o.pkg, o.doc, { report });
  assert.deepEqual(report.parts, ['/word/document.xml']);
  await checkCommentedPackage(await o.pkg.save());
});

test('a part of answers that is not XML is left alone, and the report says that answers were not written', async () => {
  const bytes = await buildSpec({ ...(await import('./doc-comments-fixtures')).commentedSpec({ body: ONE, comments: [{ id: 1, text: 'x', paraId: '0000000A' }] }), parts: [['word/commentsExtended.xml', '<broken']], rels: [['rIdCex', EXTENDED_REL, 'commentsExtended.xml']] });
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  assert.ok(doc.warnings.some((warning) => /commentsExtended/.test(warning)), 'reading said nothing of the broken part');
  const history = new DocHistory<DocRange | undefined>(doc);
  run(history, resolveComment(doc, 1, true));
  const report = emptyWriteReport();
  await writeDocx(pkg, doc, { report });
  assert.equal(decoder.decode(await pkg.readPart('/word/commentsExtended.xml')), '<broken');
  assert.ok(report.warnings.some((warning) => /commentsExtended/.test(warning)), 'the report says nothing of the part that was not written');
});

test('buildDocx writes threads and resolved comments of a model, with the part of answers', async () => {
  const o = await open(THREAD);
  run(o.history, resolveComment(o.doc, 3, true));
  run(o.history, replyToComment(o.doc, 3, 'Fine.'));
  const bytes = await convertToDocx(o.doc);
  const facts = await checkCommentedPackage(bytes);
  assert.equal(facts.ids.length, 4);
  assert.equal((facts.extended as Map<string, unknown>).size, 4);
  const again = await sameCommentsAfterReading(bytes, o.doc);
  assert.deepEqual(listComments(again).map((c) => [c.text, c.parent, c.done]), [
    ['First', undefined, false],
    ['Answer', 1, false],
    ['Other', undefined, true],
    ['Fine.', 3, true],
  ]);
  assert.equal((await OpcPackage.open(bytes)).contentTypeOf('/word/commentsExtended.xml'), EXTENDED_TYPE);
});

test('comments of an RTF file are listed, and written by buildDocx with what is added to them', async () => {
  const doc = readRtf('{\\rtf1\\ansi{\\fonttbl{\\f0 Arial;}}\\pard Text {\\*\\atnid AB}{\\*\\atnauthor Ann Bee}\\chatn{\\*\\annotation \\pard\\plain {\\chatn }A remark.\\par}after.\\par}');
  const list = listComments(doc);
  assert.deepEqual(list.map((c) => [c.author, c.initials, c.text]), [['Ann Bee', 'AB', 'A remark.']]);
  setTracking(doc, { author: 'Me', now: NOW });
  const history = new DocHistory<DocRange | undefined>(doc);
  run(history, replyToComment(doc, list[0].id, 'An answer.'));
  const bytes = await convertToDocx(doc);
  const facts = await checkCommentedPackage(bytes);
  assert.equal(facts.ids.length, 2);
  const again = await readDocx(await OpcPackage.open(bytes));
  assert.deepEqual(listComments(again).map((c) => [c.author, c.text, c.parent !== undefined]), [
    ['Ann Bee', 'A remark.', false],
    ['Me', 'An answer.', true],
  ]);
});
