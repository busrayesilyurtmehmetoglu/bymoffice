// Tests for behaviour of the document layout engine that no test pinned down:
// each of them fails when the line of code it is about is changed.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import type { ImageItem, LayoutPage } from '../src/layout/doc/types';
import { COMPAT15, at, filler, find, lay, lines, p, pageText, r, round, sectPr, tbl, tc, texts, tr } from './doc-layout-fixtures';

const PX = 9525;
const modern = (body: string) => lay({ body, settings: COMPAT15 });
const ref = (id: number) => `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;
/** A footnote of `count` lines of 16 px: "<prefix>0", "<prefix>1", ... */
const note = (id: number, count: number, prefix = 'N') => {
  let out = `<w:footnote w:id="${id}">`;
  for (let i = 0; i < count; i++) out += p(prefix + i);
  return out + '</w:footnote>';
};
const noteLines = (page: LayoutPage, id = 1) => texts(page).filter((item) => item.source && item.source.story === 'footnote:' + id);
const PICTURE =
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="x"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic>';
const RELS: [string, string, string][] = [['rIdImg', 'image', 'media/image1.png']];
const inline = (width: number, height: number) =>
  `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${width * PX}" cy="${height * PX}"/><wp:docPr id="1" name="Picture 1"/>${PICTURE}</wp:inline></w:drawing></w:r>`;
const anchored = (hFrom: string, x: number, vFrom: string, y: number, width: number, height: number) =>
  '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="${hFrom}"><wp:posOffset>${x * PX}</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="${vFrom}"><wp:posOffset>${y * PX}</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="${width * PX}" cy="${height * PX}"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="2" name="Anchored"/>${PICTURE}</wp:anchor></w:drawing></w:r>`;
const images = (page: LayoutPage) => (page.items as ImageItem[]).filter((item) => item.type === 'image');

test('doc layout: header rows repeat only when they take less than half of the column', async () => {
  const table = (headLines: number) => {
    let rows = tr(tc(filler(headLines, 'H')), '<w:tblHeader/>');
    for (let i = 0; i < 40; i++) rows += tr(tc('R' + i));
    return tbl(rows, [1500]);
  };
  // 9 lines are 144 px of the 320 px of a column: repeated on the second page.
  const low = await modern(table(9));
  assert.equal(texts(low.pages[1])[0].text, 'H0');
  assert.equal(texts(low.pages[1])[0].repeat, true);
  assert.equal(texts(low.pages[1]).filter((item) => item.repeat).length, 9);
  // 10 lines are exactly half: not less than half, so not repeated. Neither are 11.
  for (const count of [10, 11]) {
    const tall = await modern(table(count));
    assert.equal(texts(tall.pages[1])[0].text, 'R' + (20 - count), `${count} lines`);
    assert.equal(texts(tall.pages[1]).filter((item) => item.repeat).length, 0, `${count} lines`);
  }
});

test('doc layout: footnotes carried over take at most three quarters of the next column', async () => {
  // A note of 60 lines at the first line of the document: 18 lines fit under the reference.
  const { pages } = await lay({ body: p(r('Text') + ref(1)) + filler(60), footnotes: note(1, 60) });
  assert.equal(noteLines(pages[0]).length, 18);
  // On the following pages the separator (10 px) and 14 lines (224 px) are within 240 px; a 15th would not be.
  assert.equal(noteLines(pages[1]).length, 14);
  assert.equal(noteLines(pages[2]).length, 14);
  // The text has the rest: 320 - 234 px hold 5 lines.
  assert.equal(texts(pages[1]).filter((item) => item.source && item.source.story === 'body').length, 5);
  assert.equal(pages.map((page) => noteLines(page).length).reduce((a, b) => a + b, 0), 60);
});

test('doc layout: what fits in no column beside carried footnotes is placed after two columns were given up', async () => {
  // A picture of 200 px; the note leaves 86 px on every page it is carried to.
  const { pages } = await lay({ body: p(r('Text') + ref(1)) + p(inline(100, 200)) + p('after'), footnotes: note(1, 90), rels: RELS });
  const withPicture = pages.map((page) => images(page).length);
  // Page 1: the reference and 18 lines of the note. Pages 2 and 3 are given up (72 lines of the note are left, 14 a page);
  // on page 4 the picture is placed although the note leaves it no room.
  assert.deepEqual(withPicture.slice(0, 5), [0, 0, 0, 1, 0]);
  assert.equal(noteLines(pages[3]).length, 14);
  assert.equal(pages.map((page) => noteLines(page).length).reduce((a, b) => a + b, 0), 90);
  assert.equal(pageText(pages[4]).indexOf('after'), 0);
  // What is left of the note when the text has ended takes pages of its own.
  assert.equal(pages.length, 7);
  assert.deepEqual(pages.slice(4).map((page) => noteLines(page).length), [14, 14, 2]);
});

test('doc layout: a line taken back from a column takes its footnote and the rest of it back too', async () => {
  // 17 lines, then a line with a note of three lines that keeps with the next paragraph. Under the reference
  // only the first line of the note has room; the next paragraph does not fit, so both go to the next page.
  const { pages } = await lay({ body: filler(17) + p(r('A') + ref(1), '<w:keepNext/>') + p('B'), footnotes: note(1, 3) });
  assert.equal(pages.length, 2);
  assert.equal(noteLines(pages[0]).length, 0);
  assert.equal(lines(pages[0]).length, 0, 'no separator is left behind');
  // Every line of the note is there once, in order, at the foot of the second page.
  assert.deepEqual(
    noteLines(pages[1]).map((item) => [item.text, round(item.y)]),
    [
      ['N0', 360 - 48 + 12.8],
      ['N1', 360 - 32 + 12.8],
      ['N2', 360 - 16 + 12.8],
    ],
  );
  assert.equal(pageText(pages[1]).indexOf('AB'), 0);
});

test('doc layout: a line taken back from a column takes the room its drawing kept clear back too', async () => {
  // Two columns of 150 px (x = 40 and 210), 20 lines each. The last line of the first column keeps with the next
  // paragraph and anchors a drawing that lies in the second column, at the height of its paragraph.
  const drawing = anchored('page', 260, 'paragraph', 0, 50, 16);
  const body = filler(19) + p(drawing + r('A'), '<w:keepNext/>') + p('B') + filler(17, 'M') + sectPr('<w:cols w:num="2" w:space="300"/>');
  const { pages } = await lay({ body, rels: RELS });
  assert.equal(pages.length, 1);
  const page = pages[0];
  // The line moved to the second column with its drawing: it is at the top there, the text below it.
  assert.deepEqual([images(page)[0].x, images(page)[0].y], [260, 40]);
  assert.deepEqual([find(page, 'A').x, find(page, 'A').y], [210, 56 + 12.8]);
  // Where the drawing would have been had the line stayed in the first column (y = 344) nothing is kept clear:
  // the last line of the second column is text like any other.
  assert.deepEqual([find(page, 'M16').x, find(page, 'M16').y], [210, 344 + 12.8]);
});

test('doc layout: a line with a footnote needs room for the separator and the first line of the note', async () => {
  const one = note(1, 1);
  // 17 lines above: the line (16.8 px with its mark) ends at 328.8; the note needs 26 px, and 354.8 is within 360.
  const fits = await lay({ body: filler(17) + p(r('A') + ref(1)), footnotes: one });
  assert.equal(fits.pages.length, 1);
  assert.equal(round(noteLines(fits.pages[0])[0].y), 356.8);
  // 10 px lower the line ends at 338.8. It would fit with the note line (16 px) or with the separator (10 px)
  // alone, but not with both.
  const moves = await lay({ body: filler(17) + p(r('A') + ref(1), '<w:spacing w:before="150"/>'), footnotes: one });
  assert.equal(moves.pages.length, 2);
  assert.equal(texts(moves.pages[0]).length, 17);
  assert.equal(pageText(moves.pages[1]), 'AN0');
  // A second note in the column needs no separator of its own: under a first note of one line (26 px with the
  // separator) the line that ends at 313.6 has room for its note (313.6 + 16 <= 334).
  const two = note(1, 1) + note(2, 1, 'S');
  const second = await lay({ body: p(r('first') + ref(1)) + filler(15) + p(r('A') + ref(2)), footnotes: two });
  assert.equal(second.pages.length, 1);
  assert.equal(noteLines(second.pages[0], 2).length, 1);
  // One line lower it has not: 329.6 + 16 > 334.
  const third = await lay({ body: p(r('first') + ref(1)) + filler(16) + p(r('A') + ref(2)), footnotes: two });
  assert.equal(third.pages.length, 2);
  assert.equal(noteLines(third.pages[1], 2).length, 1);
});

test('doc layout: a table row keeps clear of a drawing over its whole height', async () => {
  // A drawing from y = 100 to 150, anchored in the first paragraph. Rows of 16 px start at 56.
  const drawing = anchored('column', 0, 'paragraph', 60, 100, 50);
  let rows = '';
  for (let i = 0; i < 5; i++) rows += tr(tc('R' + i));
  const { pages } = await lay({ body: p(drawing + r('first')) + tbl(rows, [3000]), settings: COMPAT15, rels: RELS });
  // The third row would start at 88, above the drawing, and end in it at 104: it goes below the drawing.
  assert.deepEqual(
    texts(pages[0]).map((item) => [item.text, item.y]),
    [
      ['first', 52.8],
      ['R0', 68.8],
      ['R1', 84.8],
      ['R2', 150 + 12.8],
      ['R3', 166 + 12.8],
      ['R4', 182 + 12.8],
    ],
  );
});

test('doc layout: the bottom padding of a cell is part of its row and holds bottom-aligned text up', async () => {
  const padded = '<w:tcMar><w:top w:w="75" w:type="dxa"/><w:bottom w:w="150" w:type="dxa"/></w:tcMar>';
  const { pages } = await modern(tbl(tr(tc('a', padded) + tc('b')) + tr(tc('c') + tc('d')), [1500, 1500]) + p('after'));
  // 5 px above, 16 px of text, 10 px below: the second row starts at 71.
  assert.deepEqual(at(pages[0]), ['a@50,57.8', 'b@150,52.8', 'c@50,83.8', 'd@150,83.8', 'after@40,99.8']);

  // Padding for every cell of the table.
  const all = '<w:tblCellMar><w:left w:w="150" w:type="dxa"/><w:bottom w:w="300" w:type="dxa"/><w:right w:w="150" w:type="dxa"/></w:tblCellMar>';
  const table = await modern(tbl(tr(tc('a')) + tr(tc('c')), [1500], all));
  assert.deepEqual(at(table.pages[0]), ['a@50,52.8', 'c@50,88.8']);

  // Bottom alignment ends above the padding; centring is within what the paddings leave.
  const height = '<w:trHeight w:val="900" w:hRule="atLeast"/>';
  const low = await modern(tbl(tr(tc('a', padded + '<w:vAlign w:val="bottom"/>') + tc('m', padded + '<w:vAlign w:val="center"/>'), height), [1500, 1500]));
  assert.equal(find(low.pages[0], 'a').y, 40 + 60 - 10 - 3.2);
  assert.equal(find(low.pages[0], 'm').y, 40 + 5 + (60 - 15 - 16) / 2 + 12.8);

  // A row divided between pages keeps the padding below what stays on the first page.
  const split = await modern(filler(17) + tbl(tr(tc(p('s1') + p('s2') + p('s3'), '<w:tcMar><w:bottom w:w="300" w:type="dxa"/></w:tcMar>')), [1500]));
  // 48 px are left: with 20 px of padding only one line fits.
  assert.equal(pageText(split.pages[0]).slice(-5), 'L16s1');
  assert.equal(pageText(split.pages[1]), 's2s3');
});

test('doc layout: footnotes are numbered in the order the text refers to them', async () => {
  const marked =
    '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>two</w:t></w:r></w:p></w:footnote>' +
    '<w:footnote w:id="5"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>five</w:t></w:r></w:p></w:footnote>';
  const { pages } = await lay({ body: p(r('a') + ref(5)) + p(r('b') + ref(2)), footnotes: marked });
  // The note with id 5 is referred to first: it is note 1, whatever its id and its place in the notes part.
  assert.deepEqual(
    texts(pages[0]).map((item) => item.text),
    ['a', '1', 'b', '2', '1', 'five', '2', 'two'],
  );
});
