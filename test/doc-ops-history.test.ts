// Operations and their inverses, undo and redo, merging of typing, and what
// an edit costs in a large document.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import type { Block, Paragraph, Table } from '../src/model/doc/types';
import { DocHistory, applyDocOp, applyDocOps, deleteBackward, insertText, splitParagraph, type DocOp, type DocPosition, type DocRange } from '../src/ops/doc/index';
import { SECT, docPicture, open, p, pos, r, tbl, tc, texts, tr } from './doc-ops-fixtures';

const para = (text: string): Paragraph => ({ type: 'paragraph', props: {}, children: text ? [{ type: 'run', props: {}, content: [{ type: 'text', text }] }] : [] });

test('replaceBlocks returns the operation that puts the old blocks back, objects and flags included', async () => {
  const doc = await open(p('one') + p('two') + p('three'));
  const [one, two, three] = doc.body.blocks;
  const inverse = applyDocOp(doc, { type: 'replaceBlocks', story: 'body', path: [], at: 1, remove: 2, insert: [para('new')] });
  assert.deepEqual(texts(doc), ['one', 'new']);
  assert.equal(doc.body.blocks[0], one);
  assert.equal(doc.body.modified, true);
  const redo = applyDocOp(doc, inverse);
  assert.deepEqual(doc.body.blocks, [one, two, three]);
  assert.equal(doc.body.blocks[1], two);
  assert.equal('modified' in doc.body, false);
  applyDocOp(doc, redo);
  assert.deepEqual(texts(doc), ['one', 'new']);
  assert.equal(doc.body.modified, true);
});

test('a change inside a table replaces the table by a copy that shares what did not change', async () => {
  const doc = await open(tbl(tr(tc('a') + tc('b')) + tr(tc('c') + tc('d')), [1000, 1000]) + p('end'));
  const table = doc.body.blocks[0] as Table;
  const inverse = applyDocOp(doc, { type: 'replaceBlocks', story: 'body', path: [0, 1, 0], at: 0, remove: 1, insert: [para('C'), para('more')] });
  const now = doc.body.blocks[0] as Table;
  assert.notEqual(now, table);
  assert.equal(now.srcStart, undefined);
  assert.equal(now.rows[0], table.rows[0], 'the other row is shared');
  assert.equal(now.rows[1].cells[1], table.rows[1].cells[1], 'the other cell is shared');
  assert.deepEqual(texts(doc), ['a', 'b', 'C', 'more', 'd', 'end']);
  assert.equal(table.rows[1].cells[0].blocks.length, 1, 'the old table is as it was');
  applyDocOp(doc, inverse);
  assert.equal(doc.body.blocks[0], table);
  assert.ok(table.srcStart !== undefined);
});

test('operations that do not fit are refused and a list of operations is applied whole or not at all', async () => {
  const doc = await open(p('one') + tbl(tr(tc('a')), [1000]) + p('end'));
  const before = docPicture(doc);
  const bad: DocOp[] = [
    { type: 'replaceBlocks', story: 'body', path: [], at: 3, remove: 1, insert: [] },
    { type: 'replaceBlocks', story: 'body', path: [], at: -1, remove: 0, insert: [] },
    { type: 'replaceBlocks', story: 'body', path: [], at: 0, remove: 1.5, insert: [] },
    { type: 'replaceBlocks', story: 'body', path: [0, 0, 0], at: 0, remove: 0, insert: [] },
    { type: 'replaceBlocks', story: 'body', path: [1, 0, 3], at: 0, remove: 0, insert: [] },
    { type: 'replaceBlocks', story: 'body', path: [1, 0, 0], at: 2, remove: 1, insert: [] },
    { type: 'replaceBlocks', story: 'body', path: [], at: 0, remove: 0, insert: [undefined as unknown as Block] },
    { type: 'replaceBlocks', story: 'nowhere', path: [], at: 0, remove: 0, insert: [] },
    { type: 'setNote', story: 'body', note: { id: 1, blocks: [] } },
    { type: 'setNumbering', nums: [[1.5, null]] },
    { type: 'setRelationship', story: 'body', id: '', relationship: null },
    { type: 'setMedia', name: '', media: null },
    { type: 'other' } as unknown as DocOp,
  ];
  for (const op of bad) {
    assert.throws(() => applyDocOp(doc, op), RangeError, JSON.stringify(op));
    assert.equal(docPicture(doc), before);
  }
  const good: DocOp = { type: 'replaceBlocks', story: 'body', path: [], at: 0, remove: 1, insert: [para('x'), para('y')] };
  assert.throws(() => applyDocOps(doc, [good, { type: 'setMedia', name: 'm', media: { bytes: new Uint8Array(1), type: 'image/png' } }, bad[1]]), RangeError);
  assert.equal(docPicture(doc), before);
  const history = new DocHistory(doc);
  assert.throws(() => history.apply([good, bad[4]], undefined, undefined), RangeError);
  assert.equal(history.canUndo, false);
  assert.equal(docPicture(doc), before);
});

test('blocks before a section mark that is not last keep its index right', async () => {
  const doc = await open({ document: undefined, body: p('one') + SECT + p('stray') });
  assert.equal(doc.body.sectionIndex, 1);
  const inverse = applyDocOp(doc, { type: 'replaceBlocks', story: 'body', path: [], at: 0, remove: 1, insert: [para('a'), para('b'), para('c')] });
  assert.equal(doc.body.sectionIndex, 3);
  const again = applyDocOp(doc, { type: 'replaceBlocks', story: 'body', path: [], at: 3, remove: 1, insert: [] });
  assert.equal(doc.body.sectionIndex, 3, 'a change after the mark does not move it');
  applyDocOp(doc, again);
  applyDocOp(doc, inverse);
  assert.equal(doc.body.sectionIndex, 1);
  assert.deepEqual(texts(doc), ['one', 'stray']);
});

function type(doc: Parameters<typeof texts>[0], history: DocHistory, at: DocPosition, text: string): DocPosition {
  let caret = at;
  for (const char of text) {
    const edit = insertText(doc, caret, char);
    history.applyEdit(edit, { anchor: caret, focus: caret });
    caret = edit.selection.focus;
  }
  return caret;
}

test('typing a word is one undo step; a space ends it, and so does moving the caret', async () => {
  const doc = await open(p('') + p('second'));
  const history = new DocHistory(doc);
  let caret = type(doc, history, pos(0, 0), 'hello brave world');
  assert.deepEqual(texts(doc), ['hello brave world', 'second']);
  assert.deepEqual(history.undo(), { meta: { anchor: pos(0, 12), focus: pos(0, 12) } });
  assert.deepEqual(texts(doc)[0], 'hello brave ');
  history.undo();
  assert.deepEqual(texts(doc)[0], 'hello ');
  assert.deepEqual(history.redo(), { meta: { anchor: pos(0, 12), focus: pos(0, 12) } });
  history.undo();
  assert.deepEqual(history.undo(), { meta: { anchor: pos(0, 0), focus: pos(0, 0) } });
  assert.deepEqual(texts(doc)[0], '');
  assert.equal(history.canUndo, false);
  assert.equal('modified' in doc.body, false);
  while (history.redo());
  assert.deepEqual(texts(doc)[0], 'hello brave world');
  // The caret was moved: what is typed now is a step of its own.
  history.seal();
  caret = type(doc, history, pos(0, 17), 'ly');
  history.undo();
  assert.deepEqual(texts(doc)[0], 'hello brave world');
  // Typing in another paragraph, Enter and Backspace are steps of their own too.
  caret = type(doc, history, pos(0, 17), 'x');
  type(doc, history, pos(1, 0), 'y');
  history.applyEdit(splitParagraph(doc, pos(1, 1)), undefined);
  type(doc, history, pos(2, 0), 'z');
  history.applyEdit(deleteBackward(doc, pos(2, 1)), undefined);
  assert.deepEqual(texts(doc), ['hello brave worldx', 'y', 'second']);
  const seen: string[] = [];
  while (history.undo()) seen.push(texts(doc).join('|'));
  assert.deepEqual(seen.slice(0, 5), ['hello brave worldx|y|zsecond', 'hello brave worldx|y|second', 'hello brave worldx|ysecond', 'hello brave worldx|second', 'hello brave world|second']);
  void caret;
});

test('a new step forgets what was undone; old steps go when there are too many', async () => {
  const doc = await open(p(''));
  const history = new DocHistory<DocRange | undefined>(doc, 3);
  for (let i = 0; i < 5; i++) {
    history.applyEdit(insertText(doc, pos(0, i), String(i)), undefined);
    history.seal();
  }
  let undone = 0;
  while (history.undo()) undone++;
  assert.equal(undone, 3);
  assert.deepEqual(texts(doc), ['01']);
  history.redo();
  history.applyEdit(insertText(doc, pos(0, 3), 'x'), undefined);
  assert.equal(history.canRedo, false);
  assert.deepEqual(texts(doc), ['012x']);
  history.clear();
  assert.equal(history.canUndo, false);
  // Nothing to do is not a step.
  assert.deepEqual(history.apply([], undefined, undefined), []);
  assert.equal(history.canUndo, false);
});

test('typing in a document of 50,000 paragraphs costs less than 2 ms a key', async (t) => {
  let body = '';
  for (let i = 0; i < 50000; i++) body += p(r('paragraph number ' + i));
  const doc = await open(body);
  const history = new DocHistory(doc);
  const first = doc.body.blocks[0];
  const last = doc.body.blocks[49999];
  let caret = pos(25000, 5);
  const times: number[] = [];
  for (let i = 0; i < 400; i++) {
    const started = process.hrtime.bigint();
    const edit = insertText(doc, caret, i % 7 === 6 ? ' ' : 'k');
    history.applyEdit(edit, undefined);
    times.push(Number(process.hrtime.bigint() - started) / 1e6);
    caret = edit.selection.focus;
  }
  times.sort((a, b) => a - b);
  const median = times[200];
  assert.ok(median < 2, `median ${median} ms`);
  assert.equal(doc.body.blocks[0], first);
  assert.equal(doc.body.blocks[49999], last);
  assert.equal(doc.body.blocks.length, 50000);
  // Enter changes the number of blocks: still far below the limit.
  const started = process.hrtime.bigint();
  for (let i = 0; i < 100; i++) history.applyEdit(splitParagraph(doc, pos(25000, 3)), undefined);
  const split = Number(process.hrtime.bigint() - started) / 1e6 / 100;
  assert.ok(split < 2, `Enter: ${split} ms`);
  while (history.undo());
  assert.equal(doc.body.blocks.length, 50000);
  assert.equal(doc.body.blocks[25000].srcStart !== undefined, true);
  assert.equal('modified' in doc.body, false);
  t.diagnostic(`typing in 50,000 paragraphs: median ${median.toFixed(4)} ms, worst ${times[399].toFixed(3)} ms a key; Enter ${split.toFixed(4)} ms`);
});

test('typing in a paragraph of a million characters stays usable', async (t) => {
  let text = '';
  for (let i = 0; i < 100000; i++) text += 'ten chars ';
  const doc = await open(p(r(text.slice(0, 500000)) + r(text.slice(500000), '<w:b/>')));
  const history = new DocHistory(doc);
  const times: number[] = [];
  let caret = pos(0, 250000);
  for (let i = 0; i < 30; i++) {
    const started = process.hrtime.bigint();
    const edit = insertText(doc, caret, 'k');
    history.applyEdit(edit, undefined);
    times.push(Number(process.hrtime.bigint() - started) / 1e6);
    caret = edit.selection.focus;
  }
  times.sort((a, b) => a - b);
  assert.equal(texts(doc)[0].length, 1000030);
  assert.ok(times[15] < 50, `median ${times[15]} ms`);
  history.undo();
  assert.equal(texts(doc)[0].length, 1000000);
  t.diagnostic(`typing in a paragraph of 1,000,000 characters: median ${times[15].toFixed(3)} ms a key`);
});
