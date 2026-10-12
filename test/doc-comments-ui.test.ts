// The parts of the comments in the document view that need no browser: the
// colours and the dates the balloon and the pages show, and what a cut of
// commented text does (the decision: it is allowed, and the comment stays).

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { deleteRange, fragmentOf, listComments, applyDocOps } from '../src/ops/doc/index';
import { cutLoses, sliceOf } from '../src/ui/doc-clipboard';
import { COMMENT_ALPHA, commentColor, commentFill, shownDate } from '../src/ui/doc-comments';
import { cEnd, cRef, cStart, openCommented, span } from './doc-comments-fixtures';
import { p, pos, r, range } from './doc-ops-fixtures';

test('the colour of a comment depends on the name of its author and on nothing else', () => {
  assert.deepEqual(commentColor('Ann'), commentColor('Ann'));
  const seen = new Set<string>();
  for (const name of ['Ann', 'Bob', 'Cid', 'Dee', 'Eve', 'Fay', 'Gus', 'Hal', 'İpek', '', 'x'.repeat(100000)]) {
    const colour = commentColor(name);
    assert.equal(colour.length, 3);
    for (const channel of colour) assert.ok(channel >= 0 && channel <= 255 && Math.floor(channel) === channel, `a channel of the colour of "${name.slice(0, 8)}" is ${channel}`);
    seen.add(colour.join(','));
  }
  assert.ok(seen.size >= 3, 'every author got the same colour');
  // Resolved is paler than open, and the thread that is shown is stronger than both.
  const alpha = (fill: string) => Number(/, ([0-9.]+)\)$/.exec(fill)![1]);
  assert.ok(alpha(commentFill('Ann', true, false)) < alpha(commentFill('Ann', false, false)), 'a resolved comment is not paler than an open one');
  assert.ok(alpha(commentFill('Ann', false, true)) > alpha(commentFill('Ann', false, false)), 'the comment that is shown is not stronger');
  assert.equal(alpha(commentFill('Ann', false, false)), COMMENT_ALPHA.open);
  assert.match(commentFill('Ann', false, false), /^rgba\(\d+, \d+, \d+, [0-9.]+\)$/);
});

test('dates of comments are shown to the minute; what is no ISO date is shown as it is', () => {
  assert.equal(shownDate('2026-10-01T09:30:15Z'), '2026-10-01 09:30');
  assert.equal(shownDate('2026-10-01T09:30:15+03:00'), '2026-10-01 09:30');
  assert.equal(shownDate('yesterday'), 'yesterday');
  assert.equal(shownDate(undefined), '');
});

test('commented text can be cut: nothing the document loses is missing from the copy, and the comment stays about the place', async () => {
  const body = p(r('One ') + cStart(1) + r('two') + cEnd(1) + cRef(1) + r(' three.')) + p('Second.');
  const { doc } = await openCommented({ body, comments: [{ id: 1, text: 'x' }] });
  const selection = range(pos(0, 2), pos(0, 9));
  const slice = sliceOf(doc, selection.anchor, selection.focus);
  assert.ok(slice, "the selection has no copy");
  assert.equal(cutLoses(doc, selection, slice), false, 'a cut of commented text is refused');
  // The copy has the text and none of the marks.
  const copied = fragmentOf(doc, selection).blocks;
  assert.equal(JSON.stringify(copied).indexOf('comment'), -1, 'the copy carries a mark of the comment');
  applyDocOps(doc, deleteRange(doc, selection).ops);
  const [comment] = listComments(doc);
  assert.deepEqual(span(comment), [[0], 2, [0], 2]);
});
