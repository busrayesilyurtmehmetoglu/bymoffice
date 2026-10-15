// Find and Replace in the text of comments: comments are searched after every other story (the
// threads in the order of their text, each with its answers), not with "main text only"; a match
// there is replaced in its place in the comment, keeping the rest of the comment (its formatting
// too); what holds anything but plain text is counted, not replaced; the comments of a "Replace
// all" are undone with the rest of it; a comment is never a tracked change; a saved file changes
// only the comment that was changed.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import type { DocModel, Paragraph, Run } from '../src/model/doc/types';
import { OpcPackage } from '../src/opc';
import {
  COMMENT_NOT_REPLACED,
  DocHistory,
  DocSearch,
  MAX_COMMENT_TEXT,
  commentOfStory,
  commentStory,
  commentTextOffset,
  compileDocFind,
  docParagraphs,
  findAllInDocument,
  findInDocument,
  listComments,
  replaceInComment,
  replaceInComments,
  replaceMatchEdit,
  replaceMatches,
  type DocEdit,
  type DocRange,
  type ReplaceMatchesOptions,
} from '../src/ops/doc/index';
import { cEnd, cRef, cStart, checkCommentedPackage, openCommented, type CommentedSpec } from './doc-comments-fixtures';
import { p, r } from './doc-ops-fixtures';

const decoder = new TextDecoder();

const FOOTNOTES =
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="2"><w:p><w:r><w:t>cat in a footnote</w:t></w:r></w:p></w:footnote>';

/** A comment whose content is given: the mark of the comment, then `runs`, in one paragraph with an id. */
const content = (runs: string, paraId: string, more = ''): string =>
  `<w:p w14:paraId="${paraId}"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>${runs}</w:p>${more}`;

/** Body text, a footnote and three comments: 5 (answered by 7) about text before that of 1. */
const SPEC: CommentedSpec = {
  body:
    p(r('cat ') + cStart(5) + r('x') + cEnd(5) + cRef(5) + cStart(7) + cEnd(7) + cRef(7) + r(' then ') + cStart(1) + r('y') + cEnd(1) + cRef(1) + r(' end') + '<w:r><w:footnoteReference w:id="2"/></w:r>') +
    p('Tail.'),
  footnotes: FOOTNOTES,
  comments: [
    { id: 1, author: 'Ann', text: 'cat one', paraId: '0000000A' },
    { id: 5, author: 'Bob', content: content(r('a ') + r('bold cat', '<w:b/>') + r(' here'), '0000000B') },
    { id: 7, author: 'Cid', text: 'cat seven cat', paraId: '0000000C' },
  ],
  extended: [['0000000A'], ['0000000B'], ['0000000C', '0000000B']],
};

async function open(spec: CommentedSpec = SPEC) {
  const opened = await openCommented(spec);
  const history = new DocHistory<DocRange | undefined>(opened.doc);
  const save = async (): Promise<Uint8Array> => {
    await writeDocx(opened.pkg, opened.doc);
    return opened.pkg.save();
  };
  return { ...opened, history, save };
}

function replaceAll(doc: DocModel, history: DocHistory<DocRange | undefined>, text: string, replacement: string, options: Partial<ReplaceMatchesOptions> = {}) {
  history.seal();
  const result = replaceMatches(doc, findAllInDocument(doc, text), replacement, { ...options, apply: (edit: DocEdit) => history.applyEdit({ ...edit, mergeKey: 'all' }, undefined) });
  history.seal();
  return result;
}

const textOf = (doc: DocModel, id: number): string => listComments(doc).filter((comment) => comment.id === id)[0].text;

const runsOf = (doc: DocModel, id: number): Run[] => {
  const comment = doc.comments!.items.get(id)!;
  return (comment.blocks[0] as Paragraph).children.filter((inline): inline is Run => inline.type === 'run');
};

async function partsOf(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const pkg = await OpcPackage.open(bytes);
  const out = new Map<string, Uint8Array>();
  for (const name of pkg.partNames()) out.set(name, await pkg.readPart(name));
  return out;
}

async function changedParts(before: Uint8Array, after: Uint8Array): Promise<string[]> {
  const a = await partsOf(before);
  const b = await partsOf(after);
  const names = new Set([...a.keys(), ...b.keys()]);
  const out: string[] = [];
  names.forEach((name) => {
    const x = a.get(name);
    const y = b.get(name);
    if (!x || !y || x.length !== y.length || x.some((byte, i) => byte !== y[i])) out.push(name);
  });
  return out.sort();
}

const commentXmlOf = (xml: string, id: number): string => {
  const start = xml.indexOf(`<w:comment w:id="${id}"`);
  assert.ok(start >= 0, `comment ${id} is in the part`);
  return xml.slice(start, xml.indexOf('</w:comment>', start) + '</w:comment>'.length);
};

test('the comments are searched last: the body, the notes, then each thread in the order of its text with its answers; not with the main text only', async () => {
  const { doc } = await open();
  const matches = findAllInDocument(doc, 'cat');
  assert.deepEqual(
    matches.map((match) => [match.range.anchor.story, match.range.anchor.offset]),
    [
      ['body', 0],
      ['footnote:2', 0],
      ['comment:5', 7],
      ['comment:7', 0],
      ['comment:7', 10],
      ['comment:1', 0],
    ],
  );
  assert.ok(matches.every((match) => match.replaceable));
  // "Main text only": no comment.
  assert.deepEqual(
    findAllInDocument(doc, 'cat', { stories: 'body' }).map((match) => match.range.anchor.story),
    ['body'],
  );
  assert.ok(docParagraphs(doc, 'body').every((ref) => commentOfStory(ref.story) === undefined));
  // A search from a match in a comment goes on to the next one, and round the end back to the body.
  const find = compileDocFind('cat')!;
  const list = docParagraphs(doc);
  const from = (story: string, offset: number) => ({ story, path: [0], offset });
  const next = (story: string, offset: number, backward = false) => {
    const search = new DocSearch(find, list, from(story, offset), backward);
    for (;;) {
      const found = search.step(1000);
      if (found !== undefined) return found && [found.range.anchor.story, found.range.anchor.offset];
    }
  };
  assert.deepEqual(next('comment:7', 3), ['comment:7', 10]);
  assert.deepEqual(next('comment:7', 13), ['comment:1', 0]);
  assert.deepEqual(next('comment:1', 3), ['body', 0]);
  assert.deepEqual(next('body', 0, true), ['comment:1', 0]);
  assert.deepEqual(next('comment:5', 7, true), ['footnote:2', 0]);
  assert.deepEqual(findInDocument(doc, 'seven')!.range.anchor.story, 'comment:7');
});

test('the story of a comment: the id after "comment:", nothing else', () => {
  assert.equal(commentStory(12), 'comment:12');
  assert.equal(commentOfStory('comment:12'), 12);
  assert.equal(commentOfStory('comment:-3'), -3);
  for (const story of ['body', 'comment:', 'comment:1x', 'comment:12345678901', 'footnote:1', 'Comment:1', undefined, 5]) assert.equal(commentOfStory(story), undefined, String(story));
});

test('where a match is in the text of a comment: its paragraphs, those of its tables too, separated by "\\n"', async () => {
  const table = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>in cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const { doc } = await open({
    body: p(cStart(1) + r('x') + cEnd(1) + cRef(1)),
    comments: [{ id: 1, content: content(r('first'), '0000000A') + table + '<w:p><w:r><w:t>last</w:t></w:r></w:p>' }],
  });
  const comment = doc.comments!.items.get(1)!;
  assert.equal(textOf(doc, 1), 'first\nin cell\nlast');
  assert.equal(commentTextOffset(comment, [0], 2), 2);
  assert.equal(commentTextOffset(comment, [1, 0, 0, 0], 3), 9);
  assert.equal(commentTextOffset(comment, [2], 4), 18);
  assert.equal(commentTextOffset(comment, [2], 5), undefined);
  assert.equal(commentTextOffset(comment, [5], 0), undefined);
  // The match in the cell is found there, and replaced there.
  const match = findAllInDocument(doc, 'cell')[0];
  assert.deepEqual([match.range.anchor.story, match.range.anchor.path, match.range.anchor.offset], ['comment:1', [1, 0, 0, 0], 3]);
  const history = new DocHistory<DocRange | undefined>(doc);
  assert.deepEqual(replaceAll(doc, history, 'cell', 'box'), { replaced: 1, skipped: 0 });
  assert.equal(textOf(doc, 1), 'first\nin box\nlast');
});

test('Replace all changes the comments with the rest, as one step of the history, and keeps the formatting of a comment', async () => {
  const { doc, history } = await open();
  const before = listComments(doc).map((comment) => comment.text);
  const result = replaceAll(doc, history, 'cat', 'dog');
  assert.deepEqual(result, { replaced: 6, skipped: 0, last: result.last });
  assert.deepEqual(
    listComments(doc).map((comment) => [comment.id, comment.text]),
    [
      [5, 'a bold dog here'],
      [7, 'dog seven dog'],
      [1, 'dog one'],
    ],
  );
  // The bold run is still bold, with the new text in it; the runs around it are as they were.
  const runs = runsOf(doc, 5);
  const bold = runs.filter((run) => run.props.bold === true);
  assert.equal(bold.length, 1);
  assert.deepEqual(bold[0].content, [{ type: 'text', text: 'bold dog' }]);
  assert.ok(runs.some((run) => run.content.some((item) => item.type === 'annotationRef')), 'the mark of the comment stays');
  // Who wrote it, when, and the thread stay.
  const seven = listComments(doc).filter((comment) => comment.id === 7)[0];
  assert.deepEqual([seven.author, seven.parent], ['Cid', 5]);
  // One step takes all of it back.
  assert.ok(history.undo());
  assert.deepEqual(
    listComments(doc).map((comment) => comment.text),
    before,
  );
  assert.equal(findAllInDocument(doc, 'dog').length, 0);
  assert.ok(!history.canUndo);
  assert.ok(history.redo());
  assert.equal(findAllInDocument(doc, 'dog').length, 6);
});

test('saved and read again: only the comments part and the document change, and in the comments part only the comment that changed; undone, the file that was opened', async () => {
  const spec: CommentedSpec = {
    body: p(r('Some ') + cStart(1) + r('text') + cEnd(1) + cRef(1) + r(' and ') + cStart(2) + r('more') + cEnd(2) + cRef(2)),
    comments: [
      { id: 1, text: 'a cat here', paraId: '0000000A' },
      { id: 2, content: content(r('nothing to change, ') + r('in bold', '<w:b/>'), '0000000B') },
    ],
  };
  const o = await open(spec);
  const result = replaceAll(o.doc, o.history, 'cat', 'dog');
  assert.equal(result.replaced, 1);
  const saved = await o.save();
  assert.deepEqual(await changedParts(o.bytes, saved), ['/word/comments.xml']);
  await checkCommentedPackage(saved);
  const before = decoder.decode((await partsOf(o.bytes)).get('/word/comments.xml')!);
  const after = decoder.decode((await partsOf(saved)).get('/word/comments.xml')!);
  assert.equal(commentXmlOf(after, 2), commentXmlOf(before, 2), 'the other comment is written as the file had it');
  assert.match(commentXmlOf(after, 1), /a dog here/);
  assert.match(commentXmlOf(after, 1), /w14:paraId="0000000A"/);
  const again = await readDocx(await OpcPackage.open(saved));
  assert.deepEqual(
    listComments(again).map((comment) => [comment.id, comment.text]),
    [
      [1, 'a dog here'],
      [2, 'nothing to change, in bold'],
    ],
  );
  // Undone and saved: the bytes of the file that was opened.
  assert.ok(o.history.undo());
  assert.deepEqual(await o.save(), o.bytes);
});

test('the formatting of the text of a comment survives a save', async () => {
  const o = await open({
    body: p(cStart(1) + r('x') + cEnd(1) + cRef(1)),
    comments: [{ id: 1, content: content(r('plain ') + r('red cat', '<w:color w:val="FF0000"/><w:i/>'), '0000000A') }],
  });
  assert.equal(replaceAll(o.doc, o.history, 'cat', 'fox').replaced, 1);
  const saved = await o.save();
  await checkCommentedPackage(saved);
  const xml = commentXmlOf(decoder.decode((await partsOf(saved)).get('/word/comments.xml')!), 1);
  assert.match(xml, /<w:r><w:rPr><w:i\/><w:color w:val="FF0000"\/><\/w:rPr><w:t[^>]*>red fox<\/w:t><\/w:r>|<w:r><w:rPr><w:color w:val="FF0000"\/><w:i\/><\/w:rPr><w:t[^>]*>red fox<\/w:t><\/w:r>/);
  assert.match(xml, /<w:annotationRef\/>/);
});

test('what is not plain text of a comment is counted and kept: a field, a tracked insertion; a match over two runs is replaced', async () => {
  const o = await open({
    body: p(cStart(1) + r('x') + cEnd(1) + cRef(1)),
    comments: [
      {
        id: 1,
        content: content(
          '<w:fldSimple w:instr=" AUTHOR "><w:r><w:t>cat field</w:t></w:r></w:fldSimple>' +
            r(' ') +
            '<w:ins w:id="90" w:author="Zed" w:date="2026-10-01T00:00:00Z"><w:r><w:t>cat inserted</w:t></w:r></w:ins>' +
            r(' c') +
            r('at split', '<w:b/>'),
          '0000000A',
        ),
      },
    ],
  });
  const matches = findAllInDocument(o.doc, 'cat');
  assert.equal(matches.length, 3);
  assert.deepEqual(
    matches.map((match) => match.replaceable),
    [false, true, true],
  );
  const result = replaceAll(o.doc, o.history, 'cat', 'dog');
  // The field result is not replaceable; the one in another author's insertion is refused when it is replaced.
  assert.deepEqual([result.replaced, result.skipped], [1, 2]);
  assert.equal(textOf(o.doc, 1), 'cat field cat inserted dog split');
  // The replacement over two runs takes the formatting of the first character, as in the body.
  const runs = runsOf(o.doc, 1);
  const last = runs.filter((run) => run.content.some((item) => item.type === 'text' && /dog/.test(item.text)))[0];
  assert.notEqual(last.props.bold, true);
  assert.throws(() => replaceMatchEdit(o.doc, findAllInDocument(o.doc, 'inserted')[0].range, 'x'), (error: Error) => error.message === COMMENT_NOT_REPLACED);
});

test('a comment is never a tracked change: where changes are tracked the body is, the comment is not', async () => {
  const o = await open({ ...SPEC, settings: '<w:trackRevisions/>' });
  const result = replaceAll(o.doc, o.history, 'cat', 'dog');
  assert.equal(result.replaced, 6);
  for (const id of [1, 5, 7]) {
    const xml = JSON.stringify(o.doc.comments!.items.get(id)!.blocks);
    assert.ok(!/"kind":"(ins|del)"/.test(xml), `comment ${id} has no tracked change`);
    assert.ok(!/cat/.test(textOf(o.doc, id)));
  }
  // The body took a tracked change.
  assert.ok(/"kind":"ins"/.test(JSON.stringify(o.doc.body.blocks)));
});

test('the comments can be skipped (a room), or only they replaced (a file that takes comments only); each counted', async () => {
  let o = await open();
  let result = replaceAll(o.doc, o.history, 'cat', 'dog', { comments: 'skip' });
  assert.deepEqual([result.replaced, result.skipped, result.inComments, result.outsideComments], [2, 4, 4, undefined]);
  assert.equal(textOf(o.doc, 1), 'cat one');
  o = await open();
  result = replaceAll(o.doc, o.history, 'cat', 'dog', { outside: 'skip' });
  assert.deepEqual([result.replaced, result.skipped, result.inComments, result.outsideComments], [4, 2, undefined, 2]);
  assert.equal(findAllInDocument(o.doc, 'cat').length, 2);
  assert.equal(result.last, undefined, 'a replacement in a comment leaves the selection where it was');
  // The limit counts only what is made.
  o = await open();
  result = replaceAll(o.doc, o.history, 'cat', 'dog', { comments: 'skip', limit: 2 });
  assert.equal(result.replaced, 2);
  o = await open();
  result = replaceAll(o.doc, o.history, 'cat', 'dog', { limit: 2 });
  assert.deepEqual(result, { replaced: 0, skipped: 0, refused: true });
});

test('the matches of every comment are one edit of the table of comments, made before those of the other stories', async () => {
  const { doc } = await open();
  const edits: DocEdit[] = [];
  const history = new DocHistory<DocRange | undefined>(doc);
  const result = replaceMatches(doc, findAllInDocument(doc, 'cat'), 'dog', {
    apply: (edit) => {
      edits.push(edit);
      history.applyEdit({ ...edit, mergeKey: 'all' }, undefined);
    },
  });
  assert.equal(result.replaced, 6);
  // One for the four matches in three comments, one each for the body and the footnote.
  assert.equal(edits.length, 3);
  assert.deepEqual(
    edits[0].ops.map((op) => op.type),
    ['setComments'],
  );
  assert.equal(edits.filter((edit) => edit.ops.some((op) => op.type === 'setComments')).length, 1);
  // A range of a comment the document does not have, and one in no comment, fail.
  const gone: DocRange = { anchor: { story: 'comment:42', path: [0], offset: 0 }, focus: { story: 'comment:42', path: [0], offset: 2 } };
  const body: DocRange = { anchor: { story: 'body', path: [0], offset: 0 }, focus: { story: 'body', path: [0], offset: 2 } };
  assert.deepEqual(replaceInComments(doc, [gone, body], 'x'), { replaced: 0, failed: 2 });
});

test('a comment that would grow beyond its longest text is not changed', async () => {
  const long = 'cat ' + 'x'.repeat(MAX_COMMENT_TEXT - 10);
  const o = await open({ body: p(cStart(1) + r('x') + cEnd(1) + cRef(1)), comments: [{ id: 1, text: long, paraId: '0000000A' }] });
  const result = replaceAll(o.doc, o.history, 'cat', 'a much longer replacement');
  assert.deepEqual([result.replaced, result.skipped], [0, 1]);
  assert.equal(textOf(o.doc, 1), long);
  assert.equal(replaceAll(o.doc, o.history, 'cat', 'dog').replaced, 1);
});

test('replaceInComment: the ranges of one comment as one edit, the last first; what is not a range of that comment fails', async () => {
  const { doc } = await open();
  const ranges = findAllInDocument(doc, 'cat')
    .filter((match) => match.range.anchor.story === 'comment:7')
    .map((match) => match.range);
  const bad: DocRange[] = [
    { anchor: { story: 'comment:1', path: [0], offset: 0 }, focus: { story: 'comment:1', path: [0], offset: 3 } },
    { anchor: { story: 'comment:7', path: [9], offset: 0 }, focus: { story: 'comment:7', path: [9], offset: 3 } },
    { anchor: { story: 'comment:7', path: [0], offset: 2 }, focus: { story: 'comment:7', path: [0], offset: 2 } },
    { anchor: { story: 'comment:7', path: [0], offset: 0 }, focus: { story: 'comment:7', path: [0], offset: 99 } },
  ];
  const selection: DocRange = { anchor: { story: 'body', path: [1], offset: 2 }, focus: { story: 'body', path: [1], offset: 2 } };
  const result = replaceInComment(doc, 7, ranges.concat(bad), 'cow', selection);
  assert.deepEqual([result.replaced, result.failed], [2, 4]);
  assert.deepEqual(result.edit!.selection, selection);
  assert.equal(result.edit!.ops.length, 1);
  assert.equal(result.edit!.ops[0].type, 'setComments');
  // Nothing was changed yet: the edit is applied by a history.
  assert.equal(textOf(doc, 7), 'cat seven cat');
  // A replacement with a paragraph end replaces nothing.
  assert.deepEqual([replaceInComment(doc, 7, ranges, 'a\nb').replaced, replaceInComment(doc, 7, ranges, 'a\nb').failed], [0, 2]);
  assert.throws(() => replaceInComment(doc, 99, ranges, 'x'));
});
