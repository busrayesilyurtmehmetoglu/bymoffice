// Third review of the document editor: lines that show nothing but deleted text have no place
// of their own. The arrow keys pass over them, a press on one puts the caret on it, and a caret
// before deleted text that ends its line stands after that text, where typed text appears.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { DocLayout } from '../src/layout/doc/layout';
import type { LayoutPage } from '../src/layout/doc/types';
import type { DocModel } from '../src/model/doc/types';
import { DocLines, type Affinity, type LineAt, type Place } from '../src/ui/doc-caret';
import { fakeMeasurer, p, r, sectPr, STYLES } from './doc-layout-fixtures';
import { readSpec } from './docx-fixtures';

async function laid(body: string): Promise<{ doc: DocModel; pages: LayoutPage[]; lines: DocLines }> {
  const doc = await readSpec({ styles: STYLES, body: body + sectPr() });
  const layout = new DocLayout(doc, fakeMeasurer, { lineMarks: true });
  layout.layout();
  return { doc, pages: layout.pages, lines: new DocLines(() => layout.pages, fakeMeasurer) };
}

const del = (text: string) => `<w:del w:id="1" w:author="Ann"><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;
// 40 characters fill a line: 8 px each in a column of 320 px.
const LIVE = 'aaaa '.repeat(8);
const GONE = 'dddd '.repeat(8);
const at = (block: number, offset: number): Place => ({ story: 'body', path: [block], offset });

/** Walks with the arrow keys from a place and returns the tops of the lines the caret is drawn on, until it stops moving. */
function walk(lines: DocLines, from: Place, direction: 1 | -1, x: number): number[] {
  const tops: number[] = [];
  let here: LineAt | undefined = lines.lineOf(from);
  let place = from;
  let affinity: Affinity = 'start';
  for (let step = 0; step < 12 && here; step++) {
    const next = lines.verticalPlace(here, direction, x);
    if (!next) break;
    place = next.place;
    affinity = next.affinity;
    here = lines.lineOf(place, affinity, next.page);
    assert.ok(here);
    assert.equal(here.line, next.line, 'the caret is drawn on the line the key went to');
    tops.push(here.line.mark.y);
  }
  return tops;
}

test('review 3: Up and Down pass over a line of deleted text in the middle of a paragraph', async () => {
  // Lines at y 40 (first), 56 (live), 72 (deleted only), 88 (after text), 104 (last).
  const { lines } = await laid(p('first para') + p(r(LIVE) + del(GONE) + r('after text')) + p('last para'));
  assert.deepEqual(walk(lines, at(2, 0), -1, 40), [88, 56, 40]);
  assert.deepEqual(walk(lines, at(0, 0), 1, 40), [56, 88, 104]);
  // Keeping x: from the middle of the live line down to the middle of the text after the deletion.
  const down = lines.verticalPlace(lines.lineOf(at(1, 5))!, 1, 40 + 5 * 8)!;
  assert.deepEqual([down.place.path[0], down.place.offset], [1, 45]);
});

test('review 3: Up and Down pass over deleted lines at the end and at the start of a paragraph', async () => {
  // The paragraph ends with two lines of deleted text: 56 live, 72 and 88 deleted, the caret at its end is on 88.
  const end = await laid(p('first para') + p(r(LIVE) + del(GONE + GONE)) + p('last para'));
  assert.deepEqual(walk(end.lines, at(2, 0), -1, 40), [88, 56, 40]);
  assert.deepEqual(walk(end.lines, at(0, 0), 1, 40), [56, 88, 104]);
  // It begins with two: 56 and 72 deleted, 88 live.
  const start = await laid(p('first para') + p(del(GONE + GONE) + r('live')) + p('last para'));
  assert.deepEqual(walk(start.lines, at(2, 0), -1, 40), [88, 40]);
  assert.deepEqual(walk(start.lines, at(0, 0), 1, 40), [88, 104]);
});

test('review 3: a press on a line of deleted text puts the caret on that line', async () => {
  const { lines } = await laid(p('first para') + p(r(LIVE) + del(GONE + GONE) + r('after text')) + p('last para'));
  // Lines: 56 live, 72 and 88 deleted, 104 "after text".
  for (const y of [72, 88]) {
    const hit = lines.hit(0, 100, y + 8)!;
    assert.deepEqual([hit.place.path[0], hit.place.offset], [1, 40]);
    const box = lines.caret(hit.place, hit.affinity, 0)!;
    assert.equal(box.line.mark.y, y, `pressed at ${y}`);
    assert.equal(box.line, hit.line);
    // After the deleted text of the line: what is typed there follows it.
    assert.equal(box.x, 40 + 40 * 8);
    // And the arrow keys leave it in both directions.
    assert.equal(lines.verticalPlace(box, -1, 100)!.line.mark.y, 56);
    assert.equal(lines.verticalPlace(box, 1, 100)!.line.mark.y, 104);
  }
  // The same place without that press is on the line the text goes on at.
  assert.equal(lines.caret(at(1, 40))!.line.mark.y, 104);
  assert.equal(lines.caret(at(1, 40), 'end')!.line.mark.y, 56);
  // A line preference that no line answers to is the start, as before.
  assert.equal(lines.caret(at(1, 40), 7)!.line.mark.y, 104);
  assert.equal(lines.caret(at(1, 3), 1)!.line.mark.y, 56);
});

test('review 3: the caret before deleted text that ends its line stands after it', async () => {
  // "keep" and, deleted, " gone": typed text goes after the deletion, and that is where the caret is.
  const one = await laid(p(r('keep') + del(' gone')));
  assert.equal(one.lines.caret(at(0, 4))!.x, 40 + 9 * 8);
  // Inside the text nothing changes.
  assert.equal(one.lines.caret(at(0, 2))!.x, 40 + 2 * 8);
  // Live text after the deletion: the caret is before that text, as it was.
  const between = await laid(p(r('keep') + del(' gone') + r(' new')));
  assert.equal(between.lines.caret(at(0, 4))!.x, 40 + 9 * 8);
  // The deletion wraps: at the end of the first line the caret is after what that line shows of it.
  const wrapped = await laid(p(r('aaaa '.repeat(4)) + del(GONE) + r('after')));
  const first = wrapped.lines.caret(at(0, 20), 'end')!;
  assert.equal(first.line.mark.y, 40);
  // (The space that ends the line takes no room.)
  assert.equal(first.x, 40 + 39 * 8);
  const second = wrapped.lines.caret(at(0, 20))!;
  assert.equal(second.line.mark.y, 56);
  assert.equal(second.x, 40 + 20 * 8);
  // Deleted text in another cell of the same row, or before the caret, does not move it.
  const before = await laid(p(del('gone ') + r('keep')));
  assert.equal(before.lines.caret(at(0, 4))!.x, 40 + 9 * 8);
  assert.equal(before.lines.caret(at(0, 0))!.x, 40 + 5 * 8);
});
