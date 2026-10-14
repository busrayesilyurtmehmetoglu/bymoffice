// Find and Replace, after review: judging matches in a paragraph full of fields stays cheap,
// the cases the first tests left open (a one-cell array formula, simple fields, a match passed
// over backwards, the way round within the paragraph a search starts in, overlapping candidates
// of whole words), what a replaced text cell keeps, and "Replace all" of documents as the view
// makes it - alone, with a limit, and in a room.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CollabDocHistory } from '../src/collab/doc/history';
import * as shared from '../src/collab/doc/commands';
import { tracing } from '../src/collab/doc/trace';
import { compileFind, findIn } from '../src/edit/find';
import { paragraphText } from '../src/model/doc/types';
import { createSheet, emptyStyleSheet, type Cell, type Workbook } from '../src/model/sheet/types';
import { DocHistory, compileDocFind, docFindCounters, docParagraphs, findAllInDocument, findInDocument, paragraphMatches, replaceMatches, type DocEdit } from '../src/ops/doc/index';
import { compileSheetFind, replaceInCell } from '../src/ops/find';
import { DocNet, assertSameDoc, bodyText, docFile, joinDoc } from './collab-doc-fixtures';
import { open, p, paragraphAt, pos, r, texts } from './doc-ops-fixtures';

const where = (found: { range: { anchor: { path: number[]; offset: number } } } | undefined) => (found ? `${found.range.anchor.path.join('.')}@${found.range.anchor.offset}` : 'none');

/** Judging every "a" of a paragraph of n letters each followed by an empty field: comparisons made. */
async function judged(n: number): Promise<{ compared: number; found: number; ms: number }> {
  let inner = '';
  for (let i = 0; i < n; i++) inner += r('a') + '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const doc = await open(p(inner));
  const before = docFindCounters.compared;
  const started = Date.now();
  const all = findAllInDocument(doc, 'a', { stories: 'body' });
  // Every "aa" has a field inside it: none is found.
  assert.equal(findInDocument(doc, 'aa', undefined, { stories: 'body' }), undefined);
  return { compared: docFindCounters.compared - before, found: all.length, ms: Date.now() - started };
}

test('find in a document: a paragraph of 60,000 fields is judged with the logarithm of what it holds, not with all of it', async () => {
  const small = await judged(1000);
  const large = await judged(60000);
  assert.equal(small.found, 1000);
  assert.equal(large.found, 60000);
  // Sixty times the matches, each judged by halving lists sixty times as long: far less than 3,600 times the work.
  const perMatch = (run: { compared: number; found: number }) => run.compared / run.found;
  assert.ok(perMatch(large) < perMatch(small) * 2, `${perMatch(small)} -> ${perMatch(large)}`);
  assert.ok(perMatch(large) < 80, String(perMatch(large)));
});

test('find in a document: counting stops at its limit inside one paragraph too', async () => {
  const doc = await open(p('a'.repeat(50000)));
  const ref = docParagraphs(doc, 'body')[0];
  assert.equal(paragraphMatches(compileDocFind('a')!, ref, undefined, 10).length, 10);
  assert.equal(paragraphMatches(compileDocFind('a')!, ref).length, 50000);
});

test('find in a document: the text of a simple field - computed or not - is found and not replaced', async () => {
  const doc = await open(p(r('page ') + '<w:fldSimple w:instr=" PAGE "><w:r><w:t>7</w:t></w:r></w:fldSimple>' + r(' see ') + '<w:fldSimple w:instr=" REF mark "><w:r><w:t>there</w:t></w:r></w:fldSimple>'));
  assert.equal(paragraphText(paragraphAt(doc, 0)), 'page 7 see there');
  const seven = findInDocument(doc, '7', undefined, { stories: 'body' });
  assert.equal(where(seven), '0@5');
  assert.equal(seven!.replaceable, false);
  const there = findInDocument(doc, 'there', undefined, { stories: 'body' });
  assert.equal(where(there), '0@11');
  assert.equal(there!.replaceable, false);
  assert.equal(findInDocument(doc, 'see', undefined, { stories: 'body' })!.replaceable, true);
});

test('find in a document: backwards, a match with something inside it is passed for the one before it', async () => {
  const doc = await open({
    body: p(r('ab ab') + '<w:r><w:footnoteReference w:id="2"/></w:r>' + r('b')),
    footnotes: '<w:footnote w:id="2"><w:p><w:r><w:t>n</w:t></w:r></w:p></w:footnote>',
  });
  assert.equal(paragraphText(paragraphAt(doc, 0)), 'ab abb');
  // "bb" runs through the note reference: never found. Backwards from the end, "ab" at 3 then at 0.
  assert.equal(findInDocument(doc, 'bb', undefined, { stories: 'body' }), undefined);
  assert.equal(where(findInDocument(doc, 'b', pos(0, 6), { stories: 'body' }, true)), '0@5');
  assert.equal(where(findInDocument(doc, 'ab', pos(0, 6), { stories: 'body' }, true)), '0@3');
  assert.equal(where(findInDocument(doc, 'ab', pos(0, 3), { stories: 'body' }, true)), '0@0');
  // Round the start of the document from the first match: back to the last one of the same paragraph.
  assert.equal(where(findInDocument(doc, 'ab', pos(0, 0), { stories: 'body' }, true)), '0@3');
  // The last candidate has the reference inside it: the one before it is found, and also the one that overlaps it.
  const two = await open({
    body: p(r('bb b') + '<w:r><w:footnoteReference w:id="2"/></w:r>' + r('bb')),
    footnotes: '<w:footnote w:id="2"><w:p><w:r><w:t>n</w:t></w:r></w:p></w:footnote>',
  });
  assert.equal(paragraphText(paragraphAt(two, 0)), 'bb bbb');
  assert.equal(where(findInDocument(two, 'bb', pos(0, 6), { stories: 'body' }, true)), '0@4');
  assert.equal(where(findInDocument(two, 'bb', pos(0, 5), { stories: 'body' }, true)), '0@0');
  assert.equal(where(findInDocument(two, 'bb', pos(0, 2), { stories: 'body' })), '0@4');
});

test('find in a document: within the paragraph it starts in, the way round takes what lies before the start only', async () => {
  const doc = await open(p('x cat y cat') + p('none') + p('cat'));
  // Forward from the middle of the first paragraph: the second "cat", the third paragraph, then round to the first "cat".
  assert.equal(where(findInDocument(doc, 'cat', pos(0, 5), { stories: 'body' })), '0@8');
  assert.equal(where(findInDocument(doc, 'cat', pos(0, 11), { stories: 'body' })), '2@0');
  assert.equal(where(findInDocument(doc, 'cat', pos(2, 3), { stories: 'body' })), '0@2');
  // A one-paragraph document: round to the match before the start, never to one after it.
  const one = await open(p('cat x cat'));
  assert.equal(where(findInDocument(one, 'cat', pos(0, 9), { stories: 'body' })), '0@0');
  assert.equal(where(findInDocument(one, 'cat', pos(0, 3), { stories: 'body' })), '0@6');
  assert.equal(where(findInDocument(one, 'cat', pos(0, 0), { stories: 'body' }, true)), '0@6');
});

test('find: of whole words, a candidate that does not stand is passed by one character, not by its length', () => {
  const find = compileFind({ text: 'a a', whole: 'word' })!;
  assert.equal(findIn(find, 'xa a a'), 3);
});

test('replace in a cell: a one-cell array formula is passed over; text keeps what else the cell carries', () => {
  const sheet = createSheet({ name: 'S', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  const cells: Cell[] = [
    { kind: 'number', value: 2, style: 0, formula: { text: 'SUM(B1:B2)', type: 'array', ref: 'A1' } },
    { kind: 'string', value: 'old text', style: 3, si: 4, ph: true, cm: 1, vm: 2 },
  ];
  sheet.rows.set(0, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells });
  const workbook: Workbook = { sheets: [sheet], styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 };
  assert.deepEqual(replaceInCell(workbook, 0, 0, 0, compileSheetFind('sum')!, 'AVERAGE'), { skipped: true, count: 1 });
  const result = replaceInCell(workbook, 0, 0, 1, compileSheetFind('old')!, 'new');
  if (!result || !('change' in result)) throw new Error('nothing was written');
  assert.deepEqual(result.change.cell, { kind: 'string', value: 'new text', style: 3, ph: true, cm: 1, vm: 2 });
});

test('replace all in a document (the view\'s way): one step, the last match first; a limit refuses it whole', async () => {
  const doc = await open(p('Cat and cat.') + p('no'));
  const history = new DocHistory(doc);
  const matches = findAllInDocument(doc, 'cat');
  const apply = (edit: DocEdit) => history.applyEdit({ ops: edit.ops, selection: edit.selection, mergeKey: 'all' }, undefined);
  assert.deepEqual(replaceMatches(doc, matches, 'bird', { apply, limit: 1 }), { replaced: 0, skipped: 0, refused: true });
  assert.deepEqual(texts(doc), ['Cat and cat.', 'no']);
  const done = replaceMatches(doc, matches, 'bird', { apply, limit: 2 });
  assert.equal(done.replaced, 2);
  assert.equal(done.skipped, 0);
  // The selection the first replacement left.
  assert.deepEqual([done.last!.focus.path, done.last!.focus.offset], [[0], 4]);
  assert.deepEqual(texts(doc), ['bird and bird.', 'no']);
  history.seal();
  history.undo();
  assert.deepEqual(texts(doc), ['Cat and cat.', 'no']);
  // A replacement the command refuses is counted, the others go on.
  let calls = 0;
  const some = replaceMatches(doc, findAllInDocument(doc, 'cat'), 'x', {
    apply,
    insert: (model, range, text) => {
      if (calls++ === 0) throw new RangeError('refused');
      return shared.insertText(model, range, text);
    },
  });
  assert.deepEqual([some.replaced, some.skipped], [1, 1]);
  assert.deepEqual(texts(doc), ['x and cat.', 'no']);
});

test('replace all in a document in a room: the existing commands travel, everybody ends with the same, and it is undone as one step', async () => {
  tracing(true);
  try {
    const bytes = await docFile(['a cat and a cat', 'cat', 'catcat']);
    const net = new DocNet();
    const a = await joinDoc(net, 'a', bytes);
    const b = await joinDoc(net, 'b', bytes);
    net.drain();
    const history = new CollabDocHistory<unknown>(a.session);
    history.seal();
    const done = replaceMatches(a.doc, findAllInDocument(a.doc, 'cat'), 'dog', {
      insert: (doc, range, text) => shared.insertText(doc, range, text),
      apply: (edit) => history.applyEdit({ ops: edit.ops, selection: edit.selection, mergeKey: 'replace-all' }, undefined),
    });
    history.seal();
    assert.equal(done.replaced, 5);
    net.drain();
    assert.equal(bodyText(a.doc), 'a dog and a dog\ndog\ndogdog');
    assertSameDoc([a, b], 'after replace all');
    assert.ok(history.undo());
    net.drain();
    assert.equal(bodyText(a.doc), 'a cat and a cat\ncat\ncatcat');
    assertSameDoc([a, b], 'after undo');
  } finally {
    tracing(false);
  }
});
