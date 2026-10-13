// Comments of a document: what the reader makes of them (threads, resolved,
// content that is more than text), the commands that add, edit, answer,
// resolve and remove them with their inverses, and what the other editing
// commands do to the marks of a comment.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { enforcedProtection, paragraphText, type DocModel, type Paragraph } from '../src/model/doc/types';
import {
  COMMENT_FIXED,
  COMMENT_GONE,
  COMMENT_NOT_HERE,
  DocHistory,
  MAX_COMMENT_TEXT,
  TRACK_UNSUPPORTED,
  acceptAll,
  addComment,
  commentAtPosition,
  deleteBackward,
  deleteRange,
  editComment,
  fragmentOf,
  initialsOf,
  insertText,
  listComments,
  listRevisions,
  rejectAll,
  removeComment,
  replyToComment,
  resolveComment,
  setRunFormat,
  setTracking,
  splitParagraph,
  type DocRange,
} from '../src/ops/doc/index';
import { COMMENT_STYLES, brief, cEnd, cRef, cStart, openCommented, span } from './doc-comments-fixtures';
import { checkInvariants, docPicture, p, pos, r, range, run, tbl, tc, texts, tr } from './doc-ops-fixtures';

const NOW = () => new Date('2026-10-12T08:09:10.500Z');

async function open(body: string, extra: Parameters<typeof openCommented>[0] extends infer S ? Partial<S> : never = {}) {
  const opened = await openCommented({ body, ...extra });
  setTracking(opened.doc, { author: 'Me Myself', now: NOW });
  return { ...opened, history: new DocHistory<DocRange | undefined>(opened.doc) };
}

const para = (doc: DocModel, index: number): Paragraph => doc.body.blocks[index] as Paragraph;
/** The inlines of a paragraph as short words: marks with their ids, runs by their text. */
const shape = (paragraph: Paragraph): string =>
  paragraph.children
    .map((inline) => {
      if (inline.type === 'commentRangeStart') return `[${inline.id}`;
      if (inline.type === 'commentRangeEnd') return `${inline.id}]`;
      if (inline.type !== 'run') return inline.type;
      return inline.content.map((item) => (item.type === 'text' ? item.text : item.type === 'commentReference' ? `*${item.id}` : item.type)).join('');
    })
    .join('|');

const ONE = p(r('One ') + cStart(1) + r('two') + cEnd(1) + cRef(1) + r(' three.')) + p('Second paragraph here.');

// ------------------------------------------------------------------ reading

test('comments are listed with their author, date, text and the text they are about', async () => {
  const { doc } = await open(ONE, { comments: [{ id: 1, author: 'Ann', text: 'Why two?', initials: 'A', date: '2026-10-01T09:00:00Z' }] });
  const list = listComments(doc);
  assert.equal(list.length, 1);
  assert.deepEqual({ ...list[0] }, {
    id: 1,
    author: 'Ann',
    initials: 'A',
    date: '2026-10-01T09:00:00Z',
    text: 'Why two?',
    done: false,
    plain: true,
    canChange: true,
    story: 'body',
    start: pos(0, 4),
    end: pos(0, 7),
  });
});

test('answers and the resolved state come from commentsExtended.xml', async () => {
  const body = p(r('One ') + cStart(1) + cStart(2) + r('two') + cEnd(1) + cRef(1) + cEnd(2) + cRef(2) + r(' three ') + cStart(3) + r('four') + cEnd(3) + cRef(3));
  const { doc } = await open(body, {
    comments: [
      { id: 1, author: 'Ann', text: 'First', paraId: '0000000A' },
      { id: 2, author: 'Bob', text: 'Answer', paraId: '0000000B' },
      { id: 3, author: 'Cid', text: 'Done one', paraId: '0000000C' },
    ],
    extended: [['0000000A'], ['0000000B', '0000000A'], ['0000000C', undefined, true]],
  });
  assert.deepEqual(brief(listComments(doc)), [
    [1, 'Ann', 'First', undefined, false],
    [2, 'Bob', 'Answer', 1, false],
    [3, 'Cid', 'Done one', undefined, true],
  ]);
  assert.equal(doc.comments!.items.get(2)!.parent, 1);
  assert.equal(doc.comments!.items.get(3)!.done, true);
  assert.equal(doc.comments!.items.get(1)!.paraId, '0000000A');
});

test('a part of answers that names nothing, answers itself or runs in a ring links nothing', async () => {
  const body = p(cStart(1) + r('a') + cEnd(1) + cRef(1) + cStart(2) + r('b') + cEnd(2) + cRef(2));
  const { doc } = await open(body, {
    comments: [
      { id: 1, text: 'One', paraId: '0000000A' },
      { id: 2, text: 'Two', paraId: '0000000B' },
    ],
    extended: [['0000000A', '0000000B'], ['0000000B', '0000000A'], ['0000000F', '0000000A'], ['0000000A', '0000000A', true]],
  });
  const list = listComments(doc);
  // Whatever the ring is cut at, no comment is its own ancestor and every comment is listed once.
  assert.deepEqual(list.map((c) => c.id).sort(), [1, 2]);
  for (const comment of list) assert.notEqual(comment.parent, comment.id, 'a comment answers itself');
  assert.ok(list.some((c) => c.parent === undefined), 'a ring of answers has no first comment');
});

test('a comment with content that is more than text is read whole and says so', async () => {
  const rich = '<w:p w14:paraId="0000000A"><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:hyperlink w:anchor="x"><w:r><w:t xml:space="preserve"> link</w:t></w:r></w:hyperlink></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p w14:paraId="0000000B"><w:r><w:t>end</w:t></w:r></w:p>';
  const { doc } = await open(ONE, { comments: [{ id: 1, content: rich }] });
  const [comment] = listComments(doc);
  assert.equal(comment.text, 'Bold link\ncell\nend');
  assert.equal(comment.plain, false);
  // Known by the id of its last paragraph.
  assert.equal(doc.comments!.items.get(1)!.paraId, '0000000B');
});

test('the place of a comment counts text as the positions of the editor do: not what is deleted, not the code of a field', async () => {
  const deleted = '<w:del w:id="9" w:author="Bob" w:date="2026-10-01T09:00:00Z"><w:r><w:delText xml:space="preserve">gone </w:delText></w:r></w:del>';
  const field = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' + r('today') + '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const { doc } = await open(p(r('One ') + deleted + field + r(' ') + cStart(1) + r('two') + cEnd(1) + cRef(1) + r(' three.')), { comments: [{ id: 1, text: 'x' }] });
  assert.equal(paragraphText(para(doc, 0)), 'One today two three.');
  const [comment] = listComments(doc);
  assert.deepEqual(span(comment), [[0], 10, [0], 13]);
  assert.equal(paragraphText(para(doc, 0)).slice(comment.start!.offset, comment.end!.offset), 'two');
});

test('a comment nothing refers to is listed last, without a place', async () => {
  const { doc } = await open(ONE, { comments: [{ id: 7, text: 'Lost' }, { id: 1, text: 'Found' }] });
  const list = listComments(doc);
  assert.deepEqual(list.map((c) => c.id), [1, 7]);
  assert.equal(list[1].start, undefined);
  assert.equal(list[1].story, undefined);
});

test('marks between blocks and a reference alone still give the comment a place', async () => {
  const body = cStart(1) + p('All of this.') + p('And this.') + cEnd(1) + p(cRef(1) + r('Tail ') + cRef(2) + r('end'));
  const { doc } = await open(body, { comments: [{ id: 1, text: 'Wide' }, { id: 2, text: 'Point' }] });
  const list = listComments(doc);
  assert.deepEqual(span(list.find((c) => c.id === 1)), [[1], 0, [2], 9]);
  assert.deepEqual(span(list.find((c) => c.id === 2)), [[4], 5, [4], 5]);
});

// ------------------------------------------------------------------- adding

test('addComment marks the range, appends the comment and undo takes both away', async () => {
  const { doc, history } = await open(p('One two three.') + p('Second.'));
  const before = docPicture(doc);
  const edit = addComment(doc, range(pos(0, 4), pos(0, 7)), 'Why two?');
  assert.equal(docPicture(doc), before, 'building the edit changed the document');
  run(history, edit);
  assert.equal(shape(para(doc, 0)), 'One |[1|two|1]|*1| three.');
  const list = listComments(doc);
  assert.deepEqual(brief(list), [[edit.id as number, 'Me Myself', 'Why two?', undefined, false]]);
  assert.equal(list[0].date, '2026-10-12T08:09:10Z');
  assert.equal(list[0].initials, 'MM');
  assert.deepEqual(span(list[0]), [[0], 4, [0], 7]);
  // The run that refers to the comment has the style Word gives it; the content is one paragraph in the style of comments.
  const ref = para(doc, 0).children[4];
  assert.equal(ref.type === 'run' ? ref.props.style : undefined, 'CommentReference');
  const content = doc.comments!.items.get(edit.id as number)!.blocks;
  assert.equal(content.length, 1);
  assert.equal((content[0] as Paragraph).props.style, 'CommentText');
  assert.equal(paragraphText(content[0] as Paragraph), 'Why two?');
  assert.deepEqual(texts(doc), ['One two three.', 'Second.']);
  checkInvariants(doc);
  history.undo();
  assert.equal(docPicture(doc), before);
  assert.equal(doc.comments, undefined, 'the collection made for the comment stayed after undo');
  history.redo();
  assert.equal(shape(para(doc, 0)), 'One |[1|two|1]|*1| three.');
  assert.equal(listComments(doc).length, 1);
});

test('addComment at a caret takes the word; between words and in an empty paragraph the place itself', async () => {
  const { doc, history } = await open(p('One two  three.') + p(''));
  const inWord = addComment(doc, pos(0, 5), 'in');
  run(history, inWord);
  assert.deepEqual(span(listComments(doc)[0]), [[0], 4, [0], 7]);
  history.undo();
  const atEnd = addComment(doc, pos(0, 7), 'end of word');
  run(history, atEnd);
  assert.deepEqual(span(listComments(doc)[0]), [[0], 4, [0], 7]);
  history.undo();
  const between = addComment(doc, pos(0, 8), 'between');
  run(history, between);
  assert.deepEqual(span(listComments(doc)[0]), [[0], 8, [0], 8]);
  assert.match(shape(para(doc, 0)), /\[\d+\|\d+\]\|\*\d+/);
  history.undo();
  run(history, addComment(doc, pos(1, 0), 'empty'));
  assert.deepEqual(span(listComments(doc)[0]), [[1], 0, [1], 0]);
  assert.equal(para(doc, 1).children.length, 3);
});

test('addComment over several paragraphs and into a table marks both ends where they are', async () => {
  const { doc, history } = await open(p('First one.') + tbl(tr(tc('Cell A') + tc('Cell B')), [2000, 2000]) + p('Last one.'));
  run(history, addComment(doc, range(pos(0, 6), pos([1, 0, 1, 0], 4)), 'wide'));
  const [comment] = listComments(doc);
  assert.deepEqual(span(comment), [[0], 6, [1, 0, 1, 0], 4]);
  assert.equal(shape(para(doc, 0)), `First |[${comment.id}|one.`);
  checkInvariants(doc);
  history.undo();
  assert.equal(listComments(doc).length, 0);
});

test('addComment takes the author, initials and date it is given, and the text as paragraphs', async () => {
  const { doc, history } = await open(p('One two.'));
  run(history, addComment(doc, range(pos(0, 0), pos(0, 3)), 'Line one\nLine\ttwo\r\nLine three', { author: 'İpek Öz', initials: 'İÖ', date: '2026-01-02T03:04:05Z' }));
  const [comment] = listComments(doc);
  assert.deepEqual([comment.author, comment.initials, comment.date, comment.text], ['İpek Öz', 'İÖ', '2026-01-02T03:04:05Z', 'Line one\nLine\ttwo\nLine three']);
  assert.equal(doc.comments!.items.get(comment.id)!.blocks.length, 3);
  assert.equal(comment.plain, true);
});

test('a comment is given the author "Author" when nobody says who is editing', async () => {
  const { doc } = await openCommented({ body: p('One two.') });
  const history = new DocHistory<DocRange | undefined>(doc);
  run(history, addComment(doc, range(pos(0, 0), pos(0, 3)), 'x'));
  assert.equal(listComments(doc)[0].author, 'Author');
  assert.equal(initialsOf('ann van der berg'), 'AVDB');
  assert.equal(initialsOf(''), '');
});

test('addComment is refused in a header, for text that is not text and for text that is too long', async () => {
  const { doc } = await open(p('One two.') + '<w:sectPr><w:headerReference w:type="default" r:id="h1"/><w:pgSz w:w="9000" w:h="9000"/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr>', {
    headers: { h1: p('Header text') },
  });
  const before = docPicture(doc);
  assert.throws(() => addComment(doc, range(pos(0, 0, 'header:h1'), pos(0, 3, 'header:h1')), 'x'), new RangeError(COMMENT_NOT_HERE));
  assert.throws(() => addComment(doc, pos(0, 1), 5 as unknown as string), RangeError);
  assert.throws(() => addComment(doc, pos(0, 1), 'x'.repeat(MAX_COMMENT_TEXT + 1)), RangeError);
  assert.throws(() => addComment(doc, pos(9, 1), 'x'), RangeError);
  assert.throws(() => addComment(doc, pos(0, 1), 'x', { date: 'yesterday' }), RangeError);
  assert.equal(docPicture(doc), before);
});

test('ids of comments are never given twice, also not after an undo', async () => {
  const { doc, history } = await open(ONE, { comments: [{ id: 1, text: 'x' }] });
  const first = addComment(doc, pos(1, 2), 'a');
  run(history, first);
  history.undo();
  const second = addComment(doc, pos(1, 2), 'b');
  assert.notEqual(second.id, first.id);
  assert.ok((second.id as number) > 1, 'a new comment took the id of a comment of the file');
});

// ------------------------------------------------------------ edit and reply

test('editComment replaces the text, keeps who wrote it, and its inverse brings the content of the file back', async () => {
  const { doc, history } = await open(ONE, { comments: [{ id: 1, author: 'Ann', text: 'Why two?', paraId: '0000000A' }] });
  const original = doc.comments!.items.get(1)!;
  assert.deepEqual(editComment(doc, 1, 'Why two?').ops, [], 'an edit that changes nothing made operations');
  run(history, editComment(doc, 1, 'Because.\nReally.'));
  const [comment] = listComments(doc);
  assert.deepEqual([comment.author, comment.text, comment.date], ['Ann', 'Because.\nReally.', '2026-10-01T09:00:00Z']);
  assert.equal(doc.comments!.items.get(1)!.paraId, '0000000A', 'the comment lost the id it is known by');
  assert.equal(doc.comments!.items.get(1)!.srcStart, undefined, 'an edited comment still points into the source');
  history.undo();
  assert.equal(doc.comments!.items.get(1), original, 'undo did not put the comment of the file back');
  assert.equal(doc.comments!.modified, undefined);
  assert.throws(() => editComment(doc, 99, 'x'), new RangeError(COMMENT_GONE));
});

test('replyToComment adds a comment of the thread with marks beside those of the first', async () => {
  const { doc, history } = await open(ONE, { comments: [{ id: 1, author: 'Ann', text: 'Why two?' }] });
  const reply = replyToComment(doc, 1, 'Because.');
  run(history, reply);
  const id = reply.id as number;
  assert.equal(shape(para(doc, 0)), `One |[1|[${id}|two|1]|${id}]|*1|*${id}| three.`);
  assert.deepEqual(brief(listComments(doc)), [
    [1, 'Ann', 'Why two?', undefined, false],
    [id, 'Me Myself', 'Because.', 1, false],
  ]);
  // Both are known by a paragraph id now: the part of answers names them by it.
  const ids = [doc.comments!.items.get(1)!.paraId, doc.comments!.items.get(id)!.paraId];
  assert.match(ids[0] as string, /^[0-9A-F]{8}$/);
  assert.match(ids[1] as string, /^[0-9A-F]{8}$/);
  assert.notEqual(ids[0], ids[1]);
  // An answer to an answer belongs to the same thread.
  const second = replyToComment(doc, id, 'Indeed.', { author: 'Cid' });
  run(history, second);
  assert.equal(listComments(doc)[2].parent, 1);
  history.undo();
  history.undo();
  assert.equal(shape(para(doc, 0)), 'One |[1|two|1]|*1| three.');
  assert.equal(doc.comments!.items.get(1)!.paraId, undefined);
  assert.equal(doc.comments!.modified, undefined);
});

test('the marks of answers stand in the order the answers were given', async () => {
  const { doc, history } = await open(ONE, { comments: [{ id: 1, author: 'Ann', text: 'Why two?' }] });
  const first = replyToComment(doc, 1, 'Because.');
  run(history, first);
  const a = first.id as number;
  const second = replyToComment(doc, 1, 'Indeed.');
  run(history, second);
  const b = second.id as number;
  // An answer to an answer goes to the end of the thread as well.
  const third = replyToComment(doc, a, 'Quite.');
  run(history, third);
  const c = third.id as number;
  assert.equal(shape(para(doc, 0)), `One |[1|[${a}|[${b}|[${c}|two|1]|${a}]|${b}]|${c}]|*1|*${a}|*${b}|*${c}| three.`);
  assert.deepEqual(listComments(doc).map((comment) => comment.id), [1, a, b, c]);
  checkInvariants(doc);
  history.undo();
  history.undo();
  history.undo();
  assert.equal(shape(para(doc, 0)), 'One |[1|two|1]|*1| three.');
});

test('an answer goes after the marks of the last answer wherever a file has them, and the marks of the file stay', async () => {
  // As Word writes a thread: each end with its reference.
  const word = p(r('One ') + cStart(1) + cStart(2) + r('two') + cEnd(1) + cRef(1) + cEnd(2) + cRef(2) + r(' three.'));
  const comments = [
    { id: 1, author: 'Ann', text: 'Why two?', paraId: '00000001' },
    { id: 2, author: 'Bob', text: 'Because.', paraId: '00000002' },
    { id: 3, author: 'Cid', text: 'Elsewhere.', paraId: '00000003' },
  ];
  let opened = await open(word + p(cStart(3) + r('Other') + cEnd(3) + cRef(3)), { comments, extended: [['00000001'], ['00000002', '00000001'], ['00000003']] });
  let reply = replyToComment(opened.doc, 1, 'Indeed.');
  run(opened.history, reply);
  let id = reply.id as number;
  assert.equal(shape(para(opened.doc, 0)), `One |[1|[2|[${id}|two|1]|*1|2]|${id}]|*2|*${id}| three.`);
  assert.equal(shape(para(opened.doc, 1)), '[3|Other|3]|*3');

  // A file whose marks are not in the order of its answers keeps them so; the new answer is the last of the thread in the text too.
  const mixed = p(r('One ') + cStart(1) + cStart(3) + cStart(2) + r('two') + cEnd(1) + cEnd(3) + cEnd(2) + cRef(1) + cRef(3) + cRef(2) + r(' three.'));
  opened = await open(mixed, { comments, extended: [['00000001'], ['00000002', '00000001'], ['00000003', '00000001']] });
  reply = replyToComment(opened.doc, 1, 'Indeed.');
  run(opened.history, reply);
  id = reply.id as number;
  assert.equal(shape(para(opened.doc, 0)), `One |[1|[3|[2|[${id}|two|1]|3]|2]|${id}]|*1|*3|*2|*${id}| three.`);

  // An answer whose marks are in another paragraph than those of the first comment: the new one stays beside the first.
  const apart = p(r('One ') + cStart(1) + r('two') + cEnd(1) + cRef(1) + r(' three.')) + p(cStart(2) + r('Other') + cEnd(2) + cRef(2));
  opened = await open(apart, { comments: comments.slice(0, 2), extended: [['00000001'], ['00000002', '00000001']] });
  reply = replyToComment(opened.doc, 1, 'Indeed.');
  run(opened.history, reply);
  id = reply.id as number;
  assert.equal(shape(para(opened.doc, 0)), `One |[1|[${id}|two|1]|${id}]|*1|*${id}| three.`);
  assert.equal(shape(para(opened.doc, 1)), '[2|Other|2]|*2');
});

test('a comment nothing refers to cannot be answered', async () => {
  const { doc } = await open(p('Plain.'), { comments: [{ id: 4, text: 'Lost' }] });
  assert.throws(() => replyToComment(doc, 4, 'x'), new RangeError(COMMENT_FIXED));
  assert.throws(() => replyToComment(doc, 5, 'x'), new RangeError(COMMENT_GONE));
});

// ---------------------------------------------------------- resolve and remove

test('resolveComment marks the whole thread, from any of its comments, and back', async () => {
  const body = p(cStart(1) + cStart(2) + r('two') + cEnd(1) + cRef(1) + cEnd(2) + cRef(2));
  const { doc, history } = await open(body, {
    comments: [{ id: 1, text: 'First', paraId: '0000000A' }, { id: 2, text: 'Answer', paraId: '0000000B' }],
    extended: [['0000000A'], ['0000000B', '0000000A']],
  });
  run(history, resolveComment(doc, 2, true));
  assert.deepEqual(listComments(doc).map((c) => c.done), [true, true]);
  assert.equal(doc.comments!.items.get(1)!.done, true);
  assert.deepEqual(resolveComment(doc, 1, true).ops, [], 'resolving a resolved thread made operations');
  run(history, resolveComment(doc, 1, false));
  assert.deepEqual(listComments(doc).map((c) => c.done), [false, false]);
  history.undo();
  history.undo();
  assert.deepEqual(listComments(doc).map((c) => c.done), [false, false]);
  assert.equal(doc.comments!.modified, undefined);
  assert.throws(() => resolveComment(doc, 1, 'yes' as unknown as boolean), RangeError);
});

test('removeComment takes the comment, its marks and its answers; an answer goes alone', async () => {
  const body = p(r('One ') + cStart(1) + cStart(2) + r('two') + cEnd(1) + cRef(1) + cEnd(2) + cRef(2) + r(' three ') + cStart(3) + r('four') + cEnd(3) + cRef(3));
  const { doc, history } = await open(body, {
    comments: [{ id: 1, text: 'First', paraId: '0000000A' }, { id: 2, text: 'Answer', paraId: '0000000B' }, { id: 3, text: 'Other', paraId: '0000000C' }],
    extended: [['0000000A'], ['0000000B', '0000000A'], ['0000000C']],
  });
  const before = docPicture(doc);
  run(history, removeComment(doc, 2));
  assert.equal(shape(para(doc, 0)), 'One |[1|two|1]|*1| three |[3|four|3]|*3');
  assert.deepEqual(listComments(doc).map((c) => c.id), [1, 3]);
  history.undo();
  run(history, removeComment(doc, 1));
  assert.equal(shape(para(doc, 0)), 'One |two| three |[3|four|3]|*3');
  assert.deepEqual(listComments(doc).map((c) => c.id), [3]);
  assert.deepEqual(texts(doc), ['One two three four']);
  checkInvariants(doc);
  history.undo();
  assert.equal(docPicture(doc), before);
  assert.equal(doc.body.modified, undefined);
  assert.throws(() => removeComment(doc, 42), new RangeError(COMMENT_GONE));
});

test('removeComment takes marks between blocks and keeps a run that holds more than the reference', async () => {
  const body = cStart(1) + p('All of this.') + cEnd(1) + p('<w:r><w:t>x</w:t><w:commentReference w:id="1"/><w:t>y</w:t></w:r>');
  const { doc, history } = await open(body, { comments: [{ id: 1, text: 'Wide' }] });
  const before = docPicture(doc);
  run(history, removeComment(doc, 1));
  assert.deepEqual(doc.body.blocks.map((block) => block.type), ['paragraph', 'paragraph']);
  assert.deepEqual(texts(doc), ['All of this.', 'xy']);
  assert.equal(listComments(doc).length, 0);
  history.undo();
  assert.equal(docPicture(doc), before);
});

test('a comment whose marks sit in a text box or in kept XML is listed, and refused for removal and answers', async () => {
  const boxed =
    '<w:r><w:pict><v:shape id="s1" style="width:100pt;height:50pt"><v:textbox><w:txbxContent>' +
    p(cStart(5) + r('in a box') + cEnd(5) + cRef(5)) +
    '</w:txbxContent></v:textbox></v:shape></w:pict></w:r>';
  const { doc } = await open(p(r('Before ') + boxed), { comments: [{ id: 5, text: 'Boxed' }] });
  const [comment] = listComments(doc);
  assert.equal(comment.canChange, false);
  const before = docPicture(doc);
  assert.throws(() => removeComment(doc, 5), new RangeError(COMMENT_FIXED));
  assert.throws(() => replyToComment(doc, 5, 'x'), new RangeError(COMMENT_FIXED));
  assert.equal(docPicture(doc), before);
});

// ------------------------------------------------------------ anchor upkeep

test('typing before, inside and after a comment moves its range with the text', async () => {
  const { doc, history } = await open(ONE, { comments: [{ id: 1, text: 'x' }] });
  run(history, insertText(doc, pos(0, 0), '>> '));
  assert.deepEqual(span(listComments(doc)[0]), [[0], 7, [0], 10]);
  run(history, insertText(doc, pos(0, 8), 'WW'));
  assert.deepEqual(span(listComments(doc)[0]), [[0], 7, [0], 12]);
  assert.equal(paragraphText(para(doc, 0)), '>> One tWWwo three.');
  run(history, insertText(doc, pos(0, 14), '!'));
  assert.deepEqual(span(listComments(doc)[0]), [[0], 7, [0], 12]);
});

test('deleting the text a comment is about leaves the comment, about a point', async () => {
  const { doc, history } = await open(ONE, { comments: [{ id: 1, text: 'x' }] });
  run(history, deleteRange(doc, range(pos(0, 2), pos(0, 9))));
  assert.equal(paragraphText(para(doc, 0)), 'Onhree.');
  const [comment] = listComments(doc);
  assert.deepEqual(span(comment), [[0], 2, [0], 2]);
  assert.equal(comment.canChange, true);
  // It can still be removed, and then nothing of it stays.
  run(history, removeComment(doc, 1));
  assert.equal(shape(para(doc, 0)).replace(/\|/g, ''), 'Onhree.');
});

test('splitting inside a comment and joining again keeps both ends', async () => {
  const { doc, history } = await open(ONE, { comments: [{ id: 1, text: 'x' }] });
  run(history, splitParagraph(doc, pos(0, 5)));
  assert.deepEqual(texts(doc), ['One t', 'wo three.', 'Second paragraph here.']);
  assert.deepEqual(span(listComments(doc)[0]), [[0], 4, [1], 2]);
  run(history, deleteBackward(doc, pos(1, 0)));
  assert.deepEqual(texts(doc), ['One two three.', 'Second paragraph here.']);
  assert.deepEqual(span(listComments(doc)[0]), [[0], 4, [0], 7]);
  // Deleting over the end of the paragraphs of a comment keeps its marks too.
  run(history, deleteRange(doc, range(pos(0, 5), pos(1, 6))));
  assert.deepEqual(texts(doc), ['One t paragraph here.']);
  assert.deepEqual(span(listComments(doc)[0]), [[0], 4, [0], 5]);
});

test('formatting text over a comment leaves its marks where they are', async () => {
  const { doc, history } = await open(ONE, { comments: [{ id: 1, text: 'x' }] });
  run(history, setRunFormat(doc, range(pos(0, 2), pos(0, 9)), { bold: true }));
  assert.deepEqual(span(listComments(doc)[0]), [[0], 4, [0], 7]);
  assert.match(shape(para(doc, 0)), /\[1\|two\|1\]\|\*1/);
});

test('a copy of commented text carries the text and no mark of the comment', async () => {
  const { doc } = await open(ONE, { comments: [{ id: 1, text: 'x' }] });
  const fragment = fragmentOf(doc, range(pos(0, 0), pos(0, 14)));
  const copied = fragment.blocks[0] as Paragraph;
  assert.equal(paragraphText(copied), 'One two three.');
  assert.equal(shape(copied).indexOf('[1'), -1, 'the copy carries the start of the comment');
  assert.equal(shape(copied).indexOf('*1'), -1, 'the copy carries the reference to the comment');
});

// ------------------------------------------------------ tracking, protection

test('with tracking on, comments are added, answered, resolved and removed without a revision', async () => {
  const { doc, history } = await open(ONE, { comments: [{ id: 1, text: 'x' }], settings: '<w:trackRevisions/>' });
  const count = () => listRevisions(doc).list.length;
  assert.equal(count(), 0);
  const added = addComment(doc, range(pos(1, 0), pos(1, 6)), 'new');
  run(history, added);
  run(history, replyToComment(doc, 1, 'answer'));
  run(history, resolveComment(doc, 1, true));
  run(history, editComment(doc, added.id as number, 'newer'));
  run(history, removeComment(doc, added.id as number));
  assert.equal(count(), 0, 'a command of comments made a tracked change');
  assert.deepEqual(texts(doc), ['One two three.', 'Second paragraph here.']);
  // Typing is tracked beside them, and the comment follows the tracked text.
  run(history, insertText(doc, pos(0, 5), 'ZZ'));
  assert.equal(count(), 1);
  assert.deepEqual(span(listComments(doc)[0]), [[0], 4, [0], 9]);
});

test('accepting and rejecting tracked changes around a comment keeps the comment', async () => {
  const body = p(r('One ') + cStart(1) + '<w:ins w:id="10" w:author="Bob" w:date="2026-10-01T09:00:00Z">' + r('new ') + '</w:ins>' + r('two') + cEnd(1) + cRef(1) + r(' three.'));
  for (const verdict of ['accept', 'reject'] as const) {
    const { doc, history } = await open(body, { comments: [{ id: 1, text: 'x' }] });
    run(history, verdict === 'accept' ? acceptAll(doc) : rejectAll(doc));
    assert.equal(paragraphText(para(doc, 0)), verdict === 'accept' ? 'One new two three.' : 'One two three.');
    assert.deepEqual(span(listComments(doc)[0]), verdict === 'accept' ? [[0], 4, [0], 11] : [[0], 4, [0], 7], verdict);
    checkInvariants(doc);
  }
});

test('a protection for comments is what the model says it is; the commands themselves work on such a document', async () => {
  const { doc, history } = await open(ONE, { comments: [{ id: 1, text: 'x' }], settings: '<w:documentProtection w:edit="comments" w:enforcement="1"/>' });
  assert.equal(enforcedProtection(doc), 'comments');
  run(history, addComment(doc, range(pos(1, 0), pos(1, 6)), 'allowed'));
  assert.equal(listComments(doc).length, 2);
});

test('the comment at a place is the one whose text holds it, the shortest when several do', async () => {
  const body = p(cStart(1) + r('One ') + cStart(2) + r('two') + cEnd(2) + cRef(2) + r(' three.') + cEnd(1) + cRef(1)) + p('Free.');
  const { doc } = await open(body, { comments: [{ id: 1, text: 'outer' }, { id: 2, text: 'inner' }] });
  const list = listComments(doc);
  assert.equal(commentAtPosition(list, pos(0, 1))!.id, 1);
  assert.equal(commentAtPosition(list, pos(0, 5))!.id, 2);
  assert.equal(commentAtPosition(list, pos(0, 7))!.id, 2);
  assert.equal(commentAtPosition(list, pos(1, 2)), undefined);
});

test('a document without the styles of comments gets them with its first comment, and loses them when that is taken back', async () => {
  const { doc, history } = await open(p('One two.'), { styles: COMMENT_STYLES.slice(0, COMMENT_STYLES.indexOf('<w:style w:type="paragraph" w:styleId="CommentText">')) });
  const before = doc.styles;
  assert.equal(before.styles.has('CommentReference') || before.styles.has('CommentText'), false);
  const edit = addComment(doc, range(pos(0, 0), pos(0, 3)), 'x');
  run(history, edit);
  const ref = para(doc, 0).children.find((inline) => inline.type === 'run' && inline.content.some((item) => item.type === 'commentReference'));
  assert.equal(ref && ref.type === 'run' ? ref.props.style : 'none', 'CommentReference');
  assert.equal((doc.comments!.items.get(edit.id as number)!.blocks[0] as Paragraph).props.style, 'CommentText');
  history.undo();
  assert.equal(doc.styles, before);
});

test('very many comments in one thread are listed in time that grows with their number, not with its square', async () => {
  const count = 30000;
  const hex = (n: number) => ('0000000' + n.toString(16).toUpperCase()).slice(-8);
  let marks = '';
  const comments: { id: number; text: string; paraId: string }[] = [];
  const extended: [string, string?][] = [];
  for (let i = 1; i <= count; i++) {
    marks += cRef(i);
    comments.push({ id: i, text: 'c' + i, paraId: hex(i) });
    extended.push(i === 1 ? [hex(1)] : [hex(i), hex(1)]);
  }
  const { doc } = await open(p(r('Text ') + marks), { comments, extended });
  const started = Date.now();
  const list = listComments(doc);
  const elapsed = Date.now() - started;
  assert.equal(list.length, count);
  assert.equal(list[count - 1].parent, 1);
  // About ten times what an idle machine takes; the quadratic way took minutes.
  assert.ok(elapsed < 5000, `listing ${count} comments took ${elapsed} ms`);
});

test('a bridge-like host can tell cheaply that the comments did not change: their texts are the same strings', async () => {
  const { doc, history } = await open(ONE, { comments: [{ id: 1, text: 'Why two?' }] });
  const before = listComments(doc)[0].text;
  run(history, insertText(doc, pos(0, 0), 'typed '));
  assert.equal(listComments(doc)[0].text, before);
});

void TRACK_UNSUPPORTED;
