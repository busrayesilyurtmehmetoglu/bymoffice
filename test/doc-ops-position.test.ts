// Positions: how they resolve, compare and move, and that they are the
// coordinates the layout engine hands out.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { DocLayout } from '../src/layout/doc/layout';
import { paragraphText } from '../src/model/doc/types';
import {
  adjacentParagraph,
  comparePositions,
  locateOffset,
  offsetOfLocation,
  movePosition,
  nextCharOffset,
  nextWordOffset,
  orderRange,
  paragraphLength,
  previousCharOffset,
  previousWordOffset,
  resolvePosition,
  snapOffset,
  storyEnd,
  storyIds,
  storyStart,
  wordAt,
} from '../src/ops/doc/index';
import { boundary, flatten, offsetOf } from '../src/ops/doc/pieces';
import { checkLayout, measurer, open, p, paragraphAt, paragraphPaths, pos, r, randomDoc, tbl, tc, tr } from './doc-ops-fixtures';
import type { Run } from '../src/model/doc/types';

const NESTED = tbl(tr(tc('a1') + tc(p('b1') + tbl(tr(tc('n1') + tc('n2')), [500, 500]) + p('b1 end'))) + tr(tc('a2') + tc('b2')), [1000, 1000]);

test('paths lead to paragraphs, in tables and nested tables too', async () => {
  const doc = await open(p('zero') + NESTED + p('last'));
  assert.equal(paragraphText(resolvePosition(doc, pos(0, 4)).paragraph), 'zero');
  assert.equal(paragraphText(resolvePosition(doc, pos([1, 0, 1, 0], 0)).paragraph), 'b1');
  const nested = resolvePosition(doc, pos([1, 0, 1, 1, 0, 1, 0], 2));
  assert.equal(paragraphText(nested.paragraph), 'n2');
  assert.deepEqual(nested.container, [1, 0, 1, 1, 0, 1]);
  assert.equal(nested.index, 0);
  assert.equal(nested.length, 2);
  for (const bad of [pos(1, 0), pos([1, 0], 0), pos([1, 0, 1], 0), pos([1, 2, 0, 0], 0), pos([1, 0, 5, 0], 0), pos(3, 0), pos(0, 5), pos([0.5], 0), pos([1, 0, 1, 1, 0, 1, 0], 3)]) {
    assert.throws(() => resolvePosition(doc, bad), RangeError, JSON.stringify(bad));
  }
  assert.throws(() => resolvePosition(doc, { story: 'body', path: 'x' as unknown as number[], offset: 0 }), RangeError);
  assert.throws(() => resolvePosition(doc, undefined as never), RangeError);
});

test('document order goes through the cells of tables', async () => {
  const doc = await open(p('zero') + NESTED + p('last'));
  assert.deepEqual(paragraphPaths(doc), [[0], [1, 0, 0, 0], [1, 0, 1, 0], [1, 0, 1, 1, 0, 0, 0], [1, 0, 1, 1, 0, 1, 0], [1, 0, 1, 2], [1, 1, 0, 0], [1, 1, 1, 0], [2]]);
  const back: number[][] = [];
  for (let path: number[] | undefined = [2]; path; path = adjacentParagraph(doc, 'body', path, -1)) back.push(path);
  assert.deepEqual(back.reverse(), paragraphPaths(doc));
  assert.deepEqual(storyStart(doc, 'body'), pos(0, 0));
  assert.deepEqual(storyEnd(doc, 'body'), pos(2, 4));
  const order = paragraphPaths(doc).map((path) => pos(path, 0));
  for (let i = 0; i < order.length; i++) {
    for (let k = 0; k < order.length; k++) assert.equal(Math.sign(comparePositions(order[i], order[k])), Math.sign(i - k));
  }
  assert.ok(comparePositions(pos(0, 1), pos(0, 3)) < 0);
  assert.deepEqual(orderRange({ anchor: pos(2, 1), focus: pos(0, 3) }).start, pos(0, 3));
  assert.throws(() => comparePositions(pos(0, 0), pos(0, 0, 'header:x')), RangeError);
});

test('moving by characters, words and paragraphs', async () => {
  const doc = await open(p('ab  cd, é̂f 😀!') + tbl(tr(tc('x') + tc('y')), [1000, 1000]) + p(''));
  const text = paragraphText(paragraphAt(doc, 0));
  const walk = (unit: 'char' | 'word', direction: 1 | -1, from: ReturnType<typeof pos>, steps: number) => {
    const out: (string | number)[] = [];
    let at = from;
    for (let i = 0; i < steps; i++) {
      at = movePosition(doc, at, unit, direction);
      out.push(at.path.length === 1 && at.path[0] === 0 ? at.offset : at.path.join('.') + '+' + at.offset);
    }
    return out;
  };
  // The letter with its two marks and the surrogate pair are one step each.
  assert.deepEqual(walk('char', 1, pos(0, 7), 6), [8, 11, 12, 13, 15, 16]);
  assert.deepEqual(walk('char', -1, pos(0, 16), 6), [15, 13, 12, 11, 8, 7]);
  assert.deepEqual(walk('char', 1, pos(0, text.length), 4), ['1.0.0.0+0', '1.0.0.0+1', '1.0.1.0+0', '1.0.1.0+1']);
  assert.deepEqual(walk('char', 1, pos([1, 0, 1, 0], 1), 2), ['2+0', '2+0']);
  assert.deepEqual(walk('char', -1, pos([1, 0, 0, 0], 0), 2), [16, 15]);
  assert.deepEqual(walk('char', -1, pos(0, 0), 1), [0]);
  assert.deepEqual(walk('word', 1, pos(0, 0), 6), [4, 6, 8, 13, 15, 16]);
  assert.deepEqual(walk('word', -1, pos(0, 16), 6), [15, 13, 8, 6, 4, 0]);
  assert.deepEqual(movePosition(doc, pos(0, 5), 'paragraphStart', 1), pos(0, 0));
  assert.deepEqual(movePosition(doc, pos(0, 5), 'paragraphEnd', 1), pos(0, 16));
  assert.deepEqual(movePosition(doc, pos(0, 5), 'paragraph', 1), pos([1, 0, 0, 0], 0));
  assert.deepEqual(movePosition(doc, pos(0, 5), 'paragraph', -1), pos(0, 0));
  assert.deepEqual(movePosition(doc, pos([1, 0, 0, 0], 0), 'paragraph', -1), pos(0, 0));
  assert.deepEqual(movePosition(doc, pos(0, 5), 'story', 1), pos(2, 0));
  assert.deepEqual(movePosition(doc, pos(2, 0), 'story', -1), pos(0, 0));
});

test('character and word offsets', () => {
  assert.equal(nextCharOffset('a😀', 1), 3);
  assert.equal(previousCharOffset('a😀', 3), 1);
  assert.equal(nextCharOffset('\ud83d', 0), 1, 'a lone surrogate is a character');
  assert.equal(nextCharOffset('é́́x', 0), 4);
  assert.equal(previousCharOffset('é́́x', 4), 0);
  // A family: person, joiner, person.
  assert.equal(nextCharOffset('👨‍👩!', 0), 5);
  assert.equal(previousCharOffset('👨‍👩!', 5), 0);
  assert.equal(nextCharOffset('ab', 5), 2);
  assert.equal(previousCharOffset('ab', 0), 0);
  assert.equal(snapOffset('a😀b', 2), 3);
  assert.equal(snapOffset('a😀b', 1), 1);
  assert.equal(snapOffset('éx', 1), 2);
  assert.equal(snapOffset('abc', 9), 3);
  assert.equal(nextWordOffset('one, two', 0), 3);
  assert.equal(nextWordOffset('one, two', 3), 5);
  assert.equal(previousWordOffset('one, two', 5), 3);
  assert.equal(previousWordOffset("don't stop", 6), 0);
  assert.deepEqual(wordAt('one, two', 6), [5, 8]);
  assert.deepEqual(wordAt('one, two', 4), [4, 5]);
  assert.deepEqual(wordAt('', 0), [0, 0]);
  // A very long run of marks is still one character, found in one pass.
  const marks = 'a' + '́'.repeat(200000);
  assert.equal(nextCharOffset(marks, 0), marks.length);
  assert.equal(previousCharOffset(marks, marks.length), 0);
});

const field = (code: string, result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r(result)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

test('offsets count what paragraphText() counts: no field codes, no deleted text; tabs and breaks one each', async () => {
  const doc = await open(
    p(
      r('ab') +
        '<w:r><w:tab/><w:t>c</w:t><w:br/><w:noBreakHyphen/><w:softHyphen/></w:r>' +
        field('REF x', 'res') +
        '<w:del w:id="1" w:author="A"><w:r><w:delText>GONE</w:delText></w:r></w:del>' +
        '<w:bookmarkStart w:id="1" w:name="b"/>' +
        r('d') +
        '<w:bookmarkEnd w:id="1"/><m:oMath><m:r><m:t>x+y</m:t></m:r></m:oMath>' +
        r('e'),
    ),
  );
  const paragraph = paragraphAt(doc, 0);
  const text = paragraphText(paragraph);
  assert.equal(text, 'ab\tc\n‑resdx+ye');
  assert.equal(paragraphLength(paragraph), text.length);
  const pieces = flatten(paragraph);
  let total = 0;
  for (const piece of pieces) total += piece.len;
  assert.equal(total, text.length);
  // The place an offset names: the piece that begins there, after what has no width, but not into a field or bookmark.
  const at = (offset: number, mode: 'caret' | 'before' | 'after', lands = offset) => {
    const list = flatten(paragraph);
    const index = boundary(list, offset, mode);
    const piece = list[index];
    assert.equal(offsetOf(list, index), lands);
    return piece ? (piece.item ? (piece.item.type === 'text' ? piece.item.text : piece.item.type === 'fieldChar' ? piece.item.charType : piece.item.type) : piece.inline!.type) : 'end';
  };
  assert.equal(at(0, 'caret'), 'ab');
  assert.equal(at(1, 'caret'), 'b');
  assert.equal(at(2, 'caret'), 'tab');
  assert.equal(at(5, 'caret'), 'noBreakHyphen');
  assert.equal(at(6, 'caret'), 'begin', 'the caret after the hyphen is after the soft hyphen and before the field');
  assert.equal(at(6, 'after'), 'softHyphen');
  assert.equal(at(6, 'before'), 'res');
  assert.equal(at(9, 'caret'), 'bookmarkStart', 'after the end of the field and the deleted run, before the bookmark');
  assert.equal(at(9, 'after'), 'end');
  assert.equal(at(10, 'caret'), 'math');
  assert.equal(at(12, 'caret', 13), 'e', 'a formula is not cut');
  assert.equal(at(12, 'before', 10), 'math');
  assert.equal(at(14, 'caret'), 'end');
});

test('every source position the layout engine hands out resolves to the text it shows', async () => {
  let checked = 0;
  let boxes = 0;
  for (let seed = 100; seed < 160; seed++) {
    const doc = await randomDoc(seed);
    checked += checkLayout(doc);
    const layout = new DocLayout(doc, measurer);
    layout.layout();
    const stories = new Set(storyIds(doc));
    for (const page of layout.pages) {
      for (const item of page.items) {
        if (item.type !== 'text' || !item.source) continue;
        assert.ok(stories.has(item.source.story), item.source.story);
        if (item.source.path.indexOf(-1) >= 0) boxes++;
      }
    }
  }
  assert.ok(checked > 1000, String(checked));
  assert.ok(boxes > 0, 'text boxes were among them');
});

test('every offset of every paragraph of random documents names the item that holds its character, and back', async () => {
  let offsets = 0;
  for (let seed = 200; seed < 240; seed++) {
    const doc = await randomDoc(seed);
    for (const story of storyIds(doc)) {
      for (const path of paragraphPaths(doc, story)) {
        const paragraph = paragraphAt(doc, path, story);
        const text = paragraphText(paragraph);
        for (let offset = 0; offset <= text.length; offset++) {
          const at = locateOffset(paragraph, offset)!;
          assert.equal(offsetOfLocation(paragraph, at), offset);
          offsets++;
          if (offset === text.length) {
            assert.equal(at.inline, paragraph.children.length);
            continue;
          }
          const inline = paragraph.children[at.inline];
          if (inline.type === 'math') assert.equal(inline.text.charAt(at.char), text.charAt(offset));
          else {
            const item = (inline as Run).content[at.item];
            const shown = item.type === 'text' ? item.text.charAt(at.char) : item.type === 'tab' ? '\t' : item.type === 'break' ? '\n' : undefined;
            assert.equal(shown, text.charAt(offset), `${story} ${path.join('.')}+${offset}`);
          }
        }
        assert.equal(locateOffset(paragraph, text.length + 1), undefined);
        assert.equal(locateOffset(paragraph, -1), undefined);
      }
    }
  }
  assert.ok(offsets > 5000, String(offsets));
});
