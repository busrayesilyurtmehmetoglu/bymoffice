// What the review of the comments found, and the limits and refusals its
// mutation run showed no test to hold: threads of any depth, what a thread
// writes into the file, the limits of the commands, where a comment cannot be
// put or changed, and the paragraph ids new comments get.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { MAX_COMMENT_ENTRIES, extendedLinks, linkComments } from '../src/formats/docx/comments';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import type { Comment, DocModel, StoryCollection } from '../src/model/doc/types';
import {
  COMMENT_FIXED,
  COMMENT_NOT_HERE,
  DocHistory,
  MAX_COMMENTS,
  addComment,
  applyDocOp,
  commentAtPosition,
  editComment,
  initialsOf,
  listComments,
  removeComment,
  replyToComment,
  resolveComment,
  setTracking,
  type DocRange,
} from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import type { XmlElement } from '../src/xml';
import { EXTENDED_REL, brief, cEnd, cRef, cStart, checkCommentedPackage, commentedSpec, extendedXml, openCommented, span, type CommentedSpec } from './doc-comments-fixtures';
import { docPicture, p, pos, r, range, run } from './doc-ops-fixtures';
import { buildDocx as buildSpec } from './docx-fixtures';

const decoder = new TextDecoder();
const NOW = () => new Date('2026-10-12T08:09:10Z');
const hex = (n: number) => ('0000000' + n.toString(16).toUpperCase()).slice(-8);

async function open(spec: CommentedSpec) {
  const opened = await openCommented(spec);
  setTracking(opened.doc, { author: 'Me', now: NOW });
  const history = new DocHistory<DocRange | undefined>(opened.doc);
  const part = async (name: string): Promise<string> => {
    await writeDocx(opened.pkg, opened.doc);
    return decoder.decode(await opened.pkg.readPart(name));
  };
  return { ...opened, history, part };
}

/** [paraId, paraIdParent, done] of every entry of a commentsExtended.xml, in its order. */
const entries = (xml: string): [string, string | undefined, string | undefined][] =>
  Array.from(xml.matchAll(/<w15:commentEx ([^>]*)\/>/g)).map((m) => [/w15:paraId="([^"]*)"/.exec(m[1])![1], (/w15:paraIdParent="([^"]*)"/.exec(m[1]) || [])[1], (/w15:done="([^"]*)"/.exec(m[1]) || [])[1]]);

const ONE = p(r('One ') + cStart(1) + r('two') + cEnd(1) + cRef(1) + r(' three.')) + p('Second paragraph here.');
const THREAD: CommentedSpec = {
  body: p(r('One ') + cStart(1) + cStart(2) + r('two') + cEnd(1) + cRef(1) + cEnd(2) + cRef(2) + r(' three.')) + p('Tail.'),
  comments: [{ id: 1, author: 'Ann', text: 'First', paraId: '0000000A' }, { id: 2, author: 'Bob', text: 'Answer', paraId: '0000000B' }],
  extended: [['0000000A'], ['0000000B', '0000000A']],
};

// --------------------------------------------------------------- deep threads

/** A document in which every comment answers the one before it. */
async function chain(count: number) {
  let inner = '';
  for (let i = 1; i <= count; i++) inner += cStart(i);
  inner += r('text');
  for (let i = 1; i <= count; i++) inner += cEnd(i) + cRef(i);
  const comments = [];
  const extended: [string, string?][] = [];
  for (let i = 1; i <= count; i++) {
    comments.push({ id: i, text: 'c' + i, paraId: hex(i) });
    extended.push(i === 1 ? [hex(1)] : [hex(i), hex(i - 1)]);
  }
  return open({ body: p(inner), comments, extended });
}

test('a thread in which each comment answers the one before is one thread, however long: none is lost from the list', async () => {
  for (const count of [3, 64, 65, 66, 200]) {
    const { doc } = await chain(count);
    const list = listComments(doc);
    assert.equal(list.length, count, `a chain of ${count} comments is listed as ${list.length}`);
    assert.deepEqual(list.map((c) => c.id), Array.from({ length: count }, (_x, i) => i + 1), `the order of a chain of ${count}`);
    assert.deepEqual(Array.from(new Set(list.slice(1).map((c) => c.parent))), [1], `the answers of a chain of ${count} do not all name its first comment`);
    assert.equal(list[0].parent, undefined);
  }
});

test('removing the first comment of a long chain takes all of it; resolving from its last comment marks all of it', async () => {
  const { doc, history } = await chain(200);
  run(history, resolveComment(doc, 200, true));
  let done = 0;
  doc.comments!.items.forEach((comment) => (done += comment.done ? 1 : 0));
  assert.equal(done, 200, 'resolving a long chain left comments of it open');
  assert.ok(listComments(doc).every((c) => c.done), 'a comment of a resolved chain is listed as open');
  history.undo();
  run(history, removeComment(doc, 1));
  assert.equal(doc.comments!.items.size, 0, 'removing the first comment left answers behind');
  assert.equal(listComments(doc).length, 0);
  assert.equal(JSON.stringify(doc.body.blocks).indexOf('comment'), -1, 'marks of the removed chain stayed in the text');
  // An answer in the middle goes alone.
  history.undo();
  run(history, removeComment(doc, 100));
  assert.equal(doc.comments!.items.size, 199);
  assert.equal(listComments(doc).length, 199, 'the answers after the removed one fell out of the list');
});

test('a very long chain of answers is listed in time that grows with its length', async () => {
  const count = 20000;
  const items = new Map<number, Comment>();
  for (let i = 1; i <= count; i++) {
    const comment: Comment = { id: i, blocks: [] };
    if (i > 1) comment.parent = i - 1;
    items.set(i, comment);
  }
  const { doc } = await open({ body: ONE });
  applyDocOp(doc, { type: 'setComments', items });
  const started = Date.now();
  const list = listComments(doc);
  const elapsed = Date.now() - started;
  assert.equal(list.length, count);
  assert.equal(list[count - 1].parent, 1);
  assert.ok(elapsed < 5000, `listing a chain of ${count} took ${elapsed} ms`);
});

test('comments that answer each other in a ring are each the first of a thread, and what answers into the ring belongs to the one it names', async () => {
  // 1 -> 2 -> 3 -> 1 is a ring; 4 answers 2; 5 answers 4.
  const body = p([1, 2, 3, 4, 5].map((id) => cStart(id)).join('') + r('text') + [1, 2, 3, 4, 5].map((id) => cEnd(id) + cRef(id)).join(''));
  const { doc } = await open({
    body,
    comments: [1, 2, 3, 4, 5].map((id) => ({ id, text: 'c' + id, paraId: hex(id) })),
    extended: [[hex(1), hex(2)], [hex(2), hex(3)], [hex(3), hex(1)], [hex(4), hex(2)], [hex(5), hex(4)]],
  });
  const list = listComments(doc);
  assert.deepEqual(list.map((c) => c.id).sort(), [1, 2, 3, 4, 5]);
  assert.deepEqual(list.filter((c) => c.parent === undefined).map((c) => c.id).sort(), [1, 2, 3]);
  assert.deepEqual(list.filter((c) => c.parent !== undefined).map((c) => [c.id, c.parent]), [[4, 2], [5, 2]]);
  // The same for a model that was made by hand, ring and all.
  const items = new Map<number, Comment>([[1, { id: 1, blocks: [], parent: 2 }], [2, { id: 2, blocks: [], parent: 1 }], [3, { id: 3, blocks: [], parent: 3 }], [4, { id: 4, blocks: [], parent: 9 }]]);
  applyDocOp(doc, { type: 'setComments', items });
  assert.deepEqual(listComments(doc).map((c) => [c.id, c.parent]), [[1, undefined], [2, undefined], [3, undefined], [4, undefined]]);
});

// ------------------------------------------------------ what a thread writes

test('resolving writes w15:done="1" for every comment of the thread, and reopening "0" for every one', async () => {
  const o = await open(THREAD);
  run(o.history, resolveComment(o.doc, 1, true));
  assert.deepEqual(entries(await o.part('/word/commentsExtended.xml')), [['0000000A', undefined, '1'], ['0000000B', '0000000A', '1']]);
  run(o.history, resolveComment(o.doc, 2, false));
  assert.deepEqual(entries(await o.part('/word/commentsExtended.xml')), [['0000000A', undefined, '0'], ['0000000B', '0000000A', '0']]);
});

test('an answer to a resolved thread is written resolved, and an answer to an answer names the first comment of the thread', async () => {
  const o = await open({ ...THREAD, extended: [['0000000A', undefined, true], ['0000000B', '0000000A', true]] });
  const reply = replyToComment(o.doc, 2, 'Third');
  run(o.history, reply);
  const made = o.doc.comments!.items.get(reply.id as number)!;
  assert.equal(made.done, true, 'an answer to a resolved thread is open in the model');
  assert.equal(made.parent, 1, 'an answer to an answer names the answer, not the first comment');
  const written = entries(await o.part('/word/commentsExtended.xml'));
  assert.deepEqual(written, [['0000000A', undefined, '1'], ['0000000B', '0000000A', '1'], [made.paraId as string, '0000000A', '1']]);
  assert.deepEqual(brief(listComments(o.doc)).map((c) => [c[3], c[4]]), [[undefined, true], [1, true], [1, true]]);
});

test('in the list an answer is resolved when the first comment of its thread is, whatever the answer says itself', async () => {
  const { doc } = await open({ ...THREAD, extended: [['0000000A', undefined, true], ['0000000B', '0000000A', false]] });
  assert.equal(doc.comments!.items.get(2)!.done, undefined);
  assert.deepEqual(listComments(doc).map((c) => c.done), [true, true]);
});

test('w15:done is read as the file may spell it: "1", "true" and "on" are resolved, anything else is not', async () => {
  for (const [value, expected] of [['1', true], ['true', true], ['on', true], ['0', false], ['false', false], ['off', false], ['yes', false], ['', false]] as [string, boolean][]) {
    const extended = extendedXml([]).replace('</w15:commentsEx>', `<w15:commentEx w15:paraId="0000000A" w15:done="${value}"/></w15:commentsEx>`);
    const bytes = await buildSpec({ ...commentedSpec({ body: ONE, comments: [{ id: 1, text: 'x', paraId: '0000000A' }] }), rels: [['rIdCex', EXTENDED_REL, 'commentsExtended.xml']], parts: [['word/commentsExtended.xml', extended]] });
    const doc = await readDocx(await OpcPackage.open(bytes));
    assert.equal(listComments(doc)[0].done, expected, `w15:done="${value}"`);
  }
});

test('of two comments whose last paragraphs claim one id the first has it; the second is known by none and gets one of its own when it is needed', async () => {
  const o = await open({
    body: p(cStart(1) + r('a') + cEnd(1) + cRef(1) + cStart(2) + r('b') + cEnd(2) + cRef(2)),
    comments: [{ id: 1, text: 'One', paraId: '0000000A' }, { id: 2, text: 'Two', paraId: '0000000A' }],
    extended: [['0000000A', undefined, true]],
  });
  assert.equal(o.doc.comments!.items.get(1)!.paraId, '0000000A');
  assert.equal(o.doc.comments!.items.get(2)!.paraId, undefined, 'two comments are known by one paragraph id');
  assert.deepEqual(listComments(o.doc).map((c) => c.done), [true, false]);
  run(o.history, resolveComment(o.doc, 2, true));
  const own = o.doc.comments!.items.get(2)!.paraId as string;
  assert.notEqual(own, '0000000A');
  await writeDocx(o.pkg, o.doc);
  await checkCommentedPackage(await o.pkg.save());
});

// --------------------------------------------------------------------- limits

test('a document takes no more than MAX_COMMENTS comments: neither a new one nor an answer', async () => {
  const { doc } = await open({ body: ONE, comments: [{ id: 1, text: 'x' }] });
  const items = new Map(doc.comments!.items);
  for (let i = 0; items.size < MAX_COMMENTS; i++) items.set(1000 + i, { id: 1000 + i, blocks: [] });
  applyDocOp(doc, { type: 'setComments', items });
  assert.equal(doc.comments!.items.size, MAX_COMMENTS);
  assert.throws(() => addComment(doc, range(pos(1, 0), pos(1, 6)), 'one too many'), /no room for another comment/);
  assert.throws(() => replyToComment(doc, 1, 'one too many'), /no room for another comment/);
  assert.equal(doc.comments!.items.size, MAX_COMMENTS);
  // One below the limit, both are taken.
  items.delete(1000);
  applyDocOp(doc, { type: 'setComments', items: new Map(items) });
  assert.equal(typeof addComment(doc, range(pos(1, 0), pos(1, 6)), 'fits').id, 'number');
  assert.equal(typeof replyToComment(doc, 1, 'fits').id, 'number');
});

test('entries of commentsExtended.xml beyond MAX_COMMENT_ENTRIES are not looked at', () => {
  const children: XmlElement[] = [];
  for (let i = 1; i <= MAX_COMMENT_ENTRIES + 2; i++) children.push({ type: 'element', name: 'w15:commentEx', attrs: new Map([['w15:paraId', hex(i)], ['w15:done', '1']]), children: [] });
  const links = extendedLinks({ type: 'element', name: 'w15:commentsEx', attrs: new Map(), children });
  assert.equal(links.size, MAX_COMMENT_ENTRIES);
  assert.ok(links.has(hex(MAX_COMMENT_ENTRIES)), 'the last entry within the limit was not read');
  assert.ok(!links.has(hex(MAX_COMMENT_ENTRIES + 1)), 'an entry beyond the limit was read');
});

test('initials are the first letters of at most nine words', () => {
  assert.equal(initialsOf('a b c d e f g h i j k l'), 'ABCDEFGHI');
  assert.equal(initialsOf('one-two_three.four  five'), 'OTTFF');
});

test('setComments keeps every comment under its own id, and never takes the comments of a file away as a whole', async () => {
  const { doc } = await open({ body: ONE, comments: [{ id: 1, text: 'x' }] });
  const before = docPicture(doc);
  const one = doc.comments!.items.get(1)!;
  assert.throws(() => applyDocOp(doc, { type: 'setComments', items: new Map([[2, one]]) }), RangeError);
  assert.throws(() => applyDocOp(doc, { type: 'setComments', items: new Map([[1.5, { ...one, id: 1.5 }]]) }), RangeError);
  assert.throws(() => applyDocOp(doc, { type: 'setComments', items: new Map([[1, { id: 1 } as unknown as Comment]]) }), RangeError);
  assert.throws(() => applyDocOp(doc, { type: 'setComments', items: [] as unknown as Map<number, Comment> }), RangeError);
  assert.throws(() => applyDocOp(doc, { type: 'setComments', items: null }), RangeError);
  assert.equal(docPicture(doc), before);
  assert.notEqual(doc.comments, undefined);
});

// ------------------------------------------------------- refusals and places

const BOXED = (inner: string) => '<w:r><w:pict><v:shape id="s1" style="width:100pt;height:50pt"><v:textbox><w:txbxContent>' + inner + '</w:txbxContent></v:textbox></v:shape></w:pict></w:r>';

test('no comment is put into a text box, and one whose marks are in a text box is neither answered nor removed', async () => {
  const { doc } = await open({ body: p(r('Before ') + BOXED(p(cStart(5) + r('in a box') + cEnd(5) + cRef(5)))) + p('After.'), comments: [{ id: 5, text: 'Boxed' }] });
  const before = docPicture(doc);
  const inBox = { story: 'body', path: [0, -1, 0, 0, 0], offset: 2 };
  assert.throws(() => addComment(doc, inBox, 'x'), new RangeError(COMMENT_NOT_HERE));
  assert.throws(() => addComment(doc, { anchor: pos(1, 0), focus: inBox }, 'x'), RangeError);
  assert.throws(() => replyToComment(doc, 5, 'x'), new RangeError(COMMENT_FIXED));
  assert.throws(() => removeComment(doc, 5), new RangeError(COMMENT_FIXED));
  assert.equal(docPicture(doc), before);
  // Its text can be changed and its thread resolved: neither touches the marks.
  assert.equal(editComment(doc, 5, 'Changed').ops.length, 1);
  assert.equal(resolveComment(doc, 5, true).ops.length, 1);
});

test('a comment that begins in a text box and is referred to in the text is not answered and not removed either', async () => {
  const { doc } = await open({ body: p(r('Before ') + BOXED(p(cStart(5) + r('in a box'))) + r(' after') + cEnd(5) + cRef(5)), comments: [{ id: 5, text: 'Half boxed' }] });
  const [comment] = listComments(doc);
  assert.equal(comment.canChange, false);
  assert.notEqual(comment.start, undefined, 'the comment has no place although the text refers to it');
  const before = docPicture(doc);
  assert.throws(() => replyToComment(doc, 5, 'x'), new RangeError(COMMENT_FIXED));
  assert.throws(() => removeComment(doc, 5), new RangeError(COMMENT_FIXED));
  assert.equal(docPicture(doc), before);
});

test('marks the wrong way round are a comment about the place of the first of them', async () => {
  const { doc } = await open({ body: p(r('One ') + cEnd(1) + r('two') + cStart(1) + cRef(1) + r(' three.')), comments: [{ id: 1, text: 'Backwards' }] });
  const [comment] = listComments(doc);
  assert.deepEqual(span(comment), [[0], 7, [0], 7]);
});

test('a comment whose content sits in a wrapper of blocks is not plain', async () => {
  const wrapped = '<w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Inside a control</w:t></w:r></w:p></w:sdtContent></w:sdt>';
  const { doc } = await open({ body: ONE, comments: [{ id: 1, content: wrapped }] });
  const [comment] = listComments(doc);
  assert.equal(comment.text, 'Inside a control');
  assert.equal(comment.plain, false);
});

test('text a tracked change deleted is not counted where a comment is placed', async () => {
  const deleted = '<w:del w:id="9" w:author="Bob" w:date="2026-10-01T09:00:00Z"><w:r><w:delText xml:space="preserve">gone gone </w:delText></w:r></w:del>';
  const moved = '<w:moveFrom w:id="10" w:author="Bob" w:date="2026-10-01T09:00:00Z"><w:r><w:t xml:space="preserve">moved </w:t></w:r></w:moveFrom>';
  const { doc } = await open({ body: p(r('One ') + deleted + moved + cStart(1) + r('two') + cEnd(1) + cRef(1)), comments: [{ id: 1, text: 'x' }] });
  assert.deepEqual(span(listComments(doc)[0]), [[0], 4, [0], 7]);
});

test('the comment at a place is the innermost of those that hold it, wherever it stands in the list', async () => {
  const body = p(cStart(2) + r('One ') + cStart(1) + r('two') + cEnd(1) + cRef(1) + r(' three ') + cStart(3) + r('four') + cEnd(3) + cRef(3) + r(' five.') + cEnd(2) + cRef(2));
  const { doc } = await open({ body, comments: [{ id: 1, text: 'inner' }, { id: 2, text: 'outer' }, { id: 3, text: 'inner too' }] });
  const list = listComments(doc);
  assert.equal(commentAtPosition(list, pos(0, 5))!.id, 1);
  assert.equal(commentAtPosition(list, pos(0, 16))!.id, 3);
  assert.equal(commentAtPosition(list, pos(0, 10))!.id, 2);
  assert.equal(commentAtPosition(list.slice().reverse(), pos(0, 5))!.id, 1);
});

// --------------------------------------------------------------- paragraph ids

/** The paragraph id the first new comment of a document gets. */
async function firstNewId(spec: CommentedSpec): Promise<{ id: string; doc: DocModel }> {
  const { doc } = await open(spec);
  const edit = addComment(doc, range(pos(0, 0), pos(0, 3)), 'new');
  const comments = (edit.ops.filter((op) => op.type === 'setComments')[0] as { items: Map<number, Comment> }).items;
  return { id: comments.get(edit.id as number)!.paraId as string, doc };
}

test('the paragraph id of a new comment is none that a paragraph of any part of the document has', async () => {
  const plain = (extra: Partial<CommentedSpec> = {}, bodyId = ''): CommentedSpec => ({ body: `<w:p${bodyId ? ` w14:paraId="${bodyId}"` : ''}>${r('One two three.')}</w:p>` + p('Second.'), ...extra });
  const free = (await firstNewId(plain())).id;
  assert.match(free, /^[0-9A-F]{8}$/);
  // The same document with that very id on a paragraph of the body, of a header, of a footnote.
  assert.notEqual((await firstNewId(plain({}, free))).id, free, 'a new comment got the id of a paragraph of the body');
  // A note changes the ids the document has, and with them what a new comment would get: asked first without the clash.
  const note = (paraId: string) => `<w:footnote w:id="1"><w:p w14:paraId="${paraId}"><w:r><w:t>note</w:t></w:r></w:p></w:footnote>`;
  const noteFree = (await firstNewId(plain({ footnotes: note('0000000F') }))).id;
  assert.notEqual((await firstNewId(plain({ footnotes: note(noteFree) }))).id, noteFree, 'a new comment got the id of a paragraph of a footnote');
  const withHeader: CommentedSpec = {
    body: `<w:p>${r('One two three.')}</w:p>` + p('Second.') + '<w:sectPr><w:headerReference w:type="default" r:id="h1"/><w:pgSz w:w="9000" w:h="9000"/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr>',
  };
  const headerFree = (await firstNewId({ ...withHeader, headers: { h1: p('Header') } })).id;
  assert.notEqual((await firstNewId({ ...withHeader, headers: { h1: `<w:p w14:paraId="${headerFree}"><w:r><w:t>Header</w:t></w:r></w:p>` } })).id, headerFree, 'a new comment got the id of a paragraph of a header');
});

test('the paragraph id of a new comment is none that a comment of the file has', async () => {
  const spec = (paraId: string): CommentedSpec => ({ body: p(r('One ') + cStart(1) + r('two') + cEnd(1) + cRef(1) + r(' three.')) + p('Second.'), comments: [{ id: 1, text: 'x', paraId }] });
  const free = (await firstNewId(spec('0000000A'))).id;
  assert.notEqual(free, '0000000A');
  assert.notEqual((await firstNewId(spec(free))).id, free, 'a new comment got the paragraph id of a comment of the file');
});

test('the paragraph id of a new comment is none that any paragraph of a comment of the file has, the last or another', async () => {
  // The comment of the file has two paragraphs: it is known by the second, and the id of the first is taken all the same.
  const spec = (first: string): CommentedSpec => ({
    body: p(r('One ') + cStart(1) + r('two') + cEnd(1) + cRef(1) + r(' three.')) + p('Second.'),
    comments: [{ id: 1, content: `<w:p w14:paraId="${first}"><w:r><w:t>first</w:t></w:r></w:p><w:p w14:paraId="0000000B"><w:r><w:t>second</w:t></w:r></w:p>` }],
  });
  const free = (await firstNewId(spec('0000000A'))).id;
  assert.notEqual((await firstNewId(spec(free))).id, free, 'a new comment got the id of a paragraph inside a comment of the file');
});

test('paragraph ids of new comments are below 0x80000000, not zero, and no two are the same', async () => {
  const { doc, history } = await open({ body: p('One two three four five six seven.') });
  const seen = new Set<string>();
  for (let i = 0; i < 300; i++) {
    const edit = i % 3 === 2 ? replyToComment(doc, listComments(doc)[0].id, 'r' + i) : addComment(doc, pos(0, (i * 5) % 30), 'c' + i);
    run(history, edit);
    const id = doc.comments!.items.get(edit.id as number)!.paraId as string;
    assert.match(id, /^[0-9A-F]{8}$/);
    const value = parseInt(id, 16);
    assert.ok(value > 0 && value < 0x80000000, `the paragraph id ${id} is out of range`);
    assert.ok(!seen.has(id), `the paragraph id ${id} was given twice`);
    seen.add(id);
  }
});

test('the last paragraph of a new comment is written with its paragraph id and a text id beside it', async () => {
  const o = await open({ body: ONE });
  const edit = addComment(o.doc, range(pos(0, 0), pos(0, 3)), 'Line one\nLine two');
  run(o.history, edit);
  const xml = await o.part('/word/comments.xml');
  const id = o.doc.comments!.items.get(edit.id as number)!.paraId as string;
  assert.match(xml, new RegExp(`<w:p w14:paraId="${id}" w14:textId="[0-9A-F]{8}">`), 'the paragraph the comment is known by has no text id beside its id');
  assert.equal((xml.match(/w14:paraId=/g) || []).length, 1, 'more than the last paragraph of the comment got an id');
  const again = await readDocx(await OpcPackage.open(await o.pkg.save()));
  assert.equal(again.comments!.items.get(edit.id as number)!.paraId, id);
  assert.deepEqual(again.warnings, []);
});

test('a ring of answers in a file is cut for every comment of the ring, in time that grows with the number of comments', () => {
  const count = 50000;
  const items = new Map<number, Comment>();
  const children: XmlElement[] = [];
  for (let i = 1; i <= count; i++) {
    items.set(i, { id: i, blocks: [{ type: 'paragraph', props: {}, children: [], attrs: [['w14:paraId', hex(i)]] }] });
    // Each answers the one before; the first answers the last.
    children.push({ type: 'element', name: 'w15:commentEx', attrs: new Map([['w15:paraId', hex(i)], ['w15:paraIdParent', hex(i === 1 ? count : i - 1)]]), children: [] });
  }
  const collection: StoryCollection<Comment> = { part: '/word/comments.xml', rootName: 'w:comments', rootAttrs: [['xmlns:w14', 'http://schemas.microsoft.com/office/word/2010/wordml']], relationships: new Map(), source: '', items };
  const started = Date.now();
  linkComments(collection, { type: 'element', name: 'w15:commentsEx', attrs: new Map(), children });
  const elapsed = Date.now() - started;
  let linked = 0;
  items.forEach((comment) => (linked += comment.parent !== undefined ? 1 : 0));
  assert.equal(linked, 0, 'comments of a ring still answer each other');
  assert.ok(elapsed < 5000, `linking ${count} comments took ${elapsed} ms`);
});
