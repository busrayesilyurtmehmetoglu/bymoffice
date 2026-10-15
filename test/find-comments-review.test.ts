// Find and Replace in comments, after its review: a match that runs into a tracked change of a
// comment is not replaced (the change would go without a trace, as the replacement in a comment
// is never tracked); a comment longer than the longest text may still be made shorter; every
// match of a paragraph of a comment is replaced in one pass over it; in a room the replacement
// does not travel; the match Find went to is forgotten once the document, its selection or the
// comment is no longer what it was.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { writeDocx } from '../src/formats/docx/write';
import type { DocModel, Paragraph, Run } from '../src/model/doc/types';
import { OpcPackage } from '../src/opc';
import * as shared from '../src/collab/doc/commands';
import { tracing, wiresOf } from '../src/collab/doc/trace';
import {
  COMMENT_NOT_REPLACED,
  DocHistory,
  MAX_COMMENT_TEXT,
  commentReplaceCounters,
  findAllInDocument,
  listComments,
  removeComment,
  replaceMatchEdit,
  replaceMatches,
  type DocEdit,
  type DocRange,
} from '../src/ops/doc/index';
import { foundInComment } from '../src/ui/find-at';
import { cEnd, cRef, cStart, openCommented, type CommentedSpec } from './doc-comments-fixtures';
import { p, r } from './doc-ops-fixtures';

const decoder = new TextDecoder();

const content = (runs: string, paraId: string): string =>
  `<w:p w14:paraId="${paraId}"><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>${runs}</w:p>`;
const ins = (text: string) => `<w:ins w:id="90" w:author="Zed" w:date="2026-10-01T00:00:00Z">${r(text)}</w:ins>`;
const del = (text: string) => `<w:del w:id="91" w:author="Zed" w:date="2026-10-01T00:00:00Z"><w:r><w:delText>${text}</w:delText></w:r></w:del>`;
const BODY = p(cStart(1) + r('x') + cEnd(1) + cRef(1));

async function open(spec: CommentedSpec) {
  const opened = await openCommented(spec);
  const history = new DocHistory<DocRange | undefined>(opened.doc);
  const commentsXml = async (): Promise<string> => {
    await writeDocx(opened.pkg, opened.doc);
    return decoder.decode(await (await OpcPackage.open(await opened.pkg.save())).readPart('/word/comments.xml'));
  };
  return { ...opened, history, commentsXml };
}

function replaceAll(doc: DocModel, history: DocHistory<DocRange | undefined>, text: string, replacement: string) {
  history.seal();
  const result = replaceMatches(doc, findAllInDocument(doc, text), replacement, { apply: (edit: DocEdit) => history.applyEdit({ ...edit, mergeKey: 'all' }, undefined) });
  history.seal();
  return result;
}

const textOf = (doc: DocModel, id: number): string => listComments(doc).filter((comment) => comment.id === id)[0].text;
const blocksOf = (doc: DocModel, id: number): string => JSON.stringify(doc.comments!.items.get(id)!.blocks, (key, value) => (key === 'parent' ? undefined : value));
const runsOf = (doc: DocModel, id: number): Run[] => (doc.comments!.items.get(id)!.blocks[0] as Paragraph).children.filter((inline): inline is Run => inline.type === 'run');

test('a match that runs into a tracked change of a comment is not replaced: the change would go without a trace (also where changes are tracked)', async () => {
  for (const settings of [undefined, '<w:trackRevisions/>']) {
    for (const [label, runs, kept] of [
      ['insertion at its end', r('c') + ins('at'), true],
      ['insertion at its start', ins('ca') + r('t'), true],
      ['deletion inside', r('c') + del('x') + r('at'), true],
      ['deletion right after it', r('cat') + del('zz'), false],
      ['deletion right before it', del('zz') + r('cat'), false],
    ] as [string, string, boolean][]) {
      const o = await open({ body: BODY, comments: [{ id: 1, author: 'Bob', content: content(runs, '0000000A') }], ...(settings ? { settings } : {}) });
      const where = `${label}${settings ? ', tracked' : ''}`;
      const before = blocksOf(o.doc, 1);
      const result = replaceAll(o.doc, o.history, 'cat', 'dog');
      if (kept) {
        assert.deepEqual([result.replaced, result.skipped], [0, 1], where);
        assert.equal(textOf(o.doc, 1), 'cat', where);
        assert.equal(blocksOf(o.doc, 1), before, where);
        assert.throws(() => replaceMatchEdit(o.doc, findAllInDocument(o.doc, 'cat')[0].range, 'dog'), (error: Error) => error.message === COMMENT_NOT_REPLACED, where);
      } else {
        assert.deepEqual([result.replaced, result.skipped], [1, 0], where);
        assert.equal(textOf(o.doc, 1), 'dog', where);
        assert.match(await o.commentsXml(), /<w:del w:id="91" w:author="Zed"[^>]*><w:r><w:delText>zz<\/w:delText><\/w:r><\/w:del>/, where);
      }
    }
  }
});

test('a comment longer than the longest text a comment takes may still be made shorter, or kept as long; not longer', async () => {
  const long = 'cat ' + 'x'.repeat(MAX_COMMENT_TEXT + 100);
  let o = await open({ body: BODY, comments: [{ id: 1, text: long, paraId: '0000000A' }] });
  let result = replaceAll(o.doc, o.history, 'cat', '');
  assert.deepEqual([result.replaced, result.skipped], [1, 0]);
  assert.equal(textOf(o.doc, 1), long.slice(3));
  o = await open({ body: BODY, comments: [{ id: 1, text: long, paraId: '0000000A' }] });
  result = replaceAll(o.doc, o.history, 'cat', 'dog');
  assert.deepEqual([result.replaced, result.skipped], [1, 0], 'as long as it was');
  result = replaceAll(o.doc, o.history, 'dog', 'dogs');
  assert.deepEqual([result.replaced, result.skipped], [0, 1], 'longer');
});

test('every match of a paragraph of a comment is replaced in one pass over it: 4,000 runs, 4,000 matches', async () => {
  const runs = Array.from({ length: 4000 }, (_, i) => r('a ', i % 2 ? '<w:b/>' : '')).join('');
  const o = await open({ body: BODY, comments: [{ id: 1, content: content(runs, '0000000A') }] });
  const before = commentReplaceCounters.passes;
  const result = replaceAll(o.doc, o.history, 'a', 'bb');
  assert.deepEqual([result.replaced, result.skipped], [4000, 0]);
  assert.equal(commentReplaceCounters.passes - before, 1);
  assert.equal(textOf(o.doc, 1), 'bb '.repeat(4000));
  assert.equal(runsOf(o.doc, 1).filter((run) => run.props.bold === true).length, 2000, 'the bold runs stay bold');
  assert.ok(o.history.undo());
  assert.equal(textOf(o.doc, 1), 'a '.repeat(4000));
});

test('in one pass: a content control whose whole text is replaced by nothing stays, once; a match over two runs takes the first', async () => {
  const sdt = (text: string) => `<w:sdt><w:sdtPr><w:tag w:val="t"/></w:sdtPr><w:sdtContent>${r(text)}</w:sdtContent></w:sdt>`;
  let o = await open({ body: BODY, comments: [{ id: 1, content: content(r('a ') + sdt('catcat') + r(' cat'), '0000000A') }] });
  let result = replaceAll(o.doc, o.history, 'cat', '');
  assert.deepEqual([result.replaced, result.skipped], [3, 0]);
  assert.equal(textOf(o.doc, 1), 'a  ');
  const xml = await o.commentsXml();
  assert.equal(xml.split('<w:sdt>').length - 1, 1);
  assert.match(xml, /<w:sdtContent><w:r><w:t\s*\/><\/w:r><\/w:sdtContent>/);
  o = await open({ body: BODY, comments: [{ id: 1, content: content(r('c') + r('at c', '<w:b/>') + r('at'), '0000000A') }] });
  result = replaceAll(o.doc, o.history, 'cat', 'dog');
  assert.deepEqual([result.replaced, result.skipped], [2, 0]);
  assert.equal(textOf(o.doc, 1), 'dog dog');
  assert.deepEqual(
    runsOf(o.doc, 1)
      .filter((run) => run.content.some((item) => item.type === 'text'))
      .map((run) => [run.props.bold === true, run.content.map((item) => (item.type === 'text' ? item.text : '')).join('')]),
    [
      [false, 'dog'],
      [true, ' dog'],
    ],
  );
});

test('in a room a replacement in a comment does not travel: it is no command of the room, so its history refuses it', async () => {
  const { doc } = await open({ body: BODY, comments: [{ id: 1, text: 'cat seven', paraId: '0000000A' }] });
  tracing(true);
  try {
    const edit = shared.replaceMatchEdit(doc, findAllInDocument(doc, 'seven')[0].range, 'eight');
    assert.ok(edit.ops.length > 0);
    assert.equal(wiresOf(edit.ops), undefined);
  } finally {
    tracing(false);
  }
});

test('the match in a comment Find went to holds while the document, its selection and the comment are what they were', async () => {
  const { doc, history } = await open({ body: p(cStart(1) + r('x') + cEnd(1) + cRef(1)) + p('second'), comments: [{ id: 1, text: 'a cat here', paraId: '0000000A' }] });
  const range = findAllInDocument(doc, 'cat')[0].range;
  const selection: DocRange = { anchor: { story: 'body', path: [1], offset: 1 }, focus: { story: 'body', path: [1], offset: 3 } };
  const at = { range, changes: 4, selection };
  assert.deepEqual(foundInComment(at, doc, 4, { anchor: { ...selection.anchor }, focus: { ...selection.focus } }), range);
  // The document changed, the selection moved (an end, the other end, or it went), no selection then and one now.
  assert.equal(foundInComment(at, doc, 5, selection), undefined);
  assert.equal(foundInComment(at, doc, 4, { anchor: selection.anchor, focus: { ...selection.focus, offset: 2 } }), undefined);
  assert.equal(foundInComment(at, doc, 4, { anchor: { ...selection.anchor, path: [0] }, focus: selection.focus }), undefined);
  assert.equal(foundInComment(at, doc, 4, undefined), undefined);
  assert.equal(foundInComment({ ...at, selection: undefined }, doc, 4, selection), undefined);
  assert.deepEqual(foundInComment({ ...at, selection: undefined }, doc, 4, undefined), range);
  assert.equal(foundInComment(undefined, doc, 4, selection), undefined);
  // The comment is not there any more, or its text is shorter than the match: whatever the count of changes says.
  const shorter = { ...range, focus: { ...range.focus, offset: 99 } };
  assert.equal(foundInComment({ ...at, range: shorter }, doc, 4, selection), undefined);
  history.applyEdit(removeComment(doc, 1), undefined);
  assert.equal(foundInComment(at, doc, 4, selection), undefined);
});
