// Typing, deleting, breaking and joining paragraphs on the document model.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { paragraphText, type Paragraph, type Table } from '../src/model/doc/types';
import { DocHistory, deleteBackward, deleteForward, deleteRange, insertBlock, insertPageBreak, insertText, splitParagraph, type DocRange } from '../src/ops/doc/index';
import { SECT, docPicture, open, p, paragraphAt, pos, r, range, run, shapeAt, tbl, tc, texts, tr } from './doc-ops-fixtures';

const history = (doc: Parameters<typeof texts>[0]) => new DocHistory<DocRange | undefined>(doc);

test('typing inside a run keeps the run and its formatting', async () => {
  const doc = await open(p(r('hello', '<w:b/>')) + p('second'));
  const h = history(doc);
  const before = doc.body.blocks[0] as Paragraph;
  const props = (before.children[0] as { props: object }).props;
  const untouched = doc.body.blocks[1];
  const after = run(h, insertText(doc, pos(0, 2), 'XY'));
  assert.deepEqual(shapeAt(doc, 0), ['b:heXYllo']);
  assert.deepEqual(after.focus, pos(0, 4));
  const now = doc.body.blocks[0] as Paragraph;
  assert.notEqual(now, before);
  assert.equal((now.children[0] as { props: object }).props, props, 'the property bag is shared, not copied');
  // The writer contract: the edited block has no source span, its neighbour keeps its own, the story is modified.
  assert.equal(now.srcStart, undefined);
  assert.equal(now.srcEnd, undefined);
  assert.equal(doc.body.blocks[1], untouched);
  assert.ok(untouched.srcStart !== undefined);
  assert.equal(doc.body.modified, true);
  // The old paragraph was not changed in place.
  assert.equal(paragraphText(before), 'hello');
  h.undo();
  assert.equal(doc.body.blocks[0], before, 'undo puts the very object back');
  assert.ok(before.srcStart !== undefined);
  assert.equal(doc.body.modified, undefined);
  assert.equal('modified' in doc.body, false);
  h.redo();
  assert.deepEqual(texts(doc), ['heXYllo', 'second']);
  assert.equal(doc.body.modified, true);
});

test('typed text takes the formatting to its left, to its right at the start, of the mark in an empty paragraph', async () => {
  const doc = await open(p(r('bold', '<w:b/>') + r('plain') + r('it', '<w:i/>')) + p('', '<w:rPr><w:i/></w:rPr>') + p(''));
  const h = history(doc);
  run(h, insertText(doc, pos(0, 4), 'X'));
  assert.deepEqual(shapeAt(doc, 0), ['b:boldX', ':plain', 'i:it']);
  run(h, insertText(doc, pos(0, 0), 'Y'));
  assert.deepEqual(shapeAt(doc, 0), ['b:YboldX', ':plain', 'i:it']);
  run(h, insertText(doc, pos(0, 11), 'Z'));
  assert.deepEqual(shapeAt(doc, 0), ['b:YboldX', ':plainZ', 'i:it']);
  run(h, insertText(doc, pos(1, 0), 'mark'));
  assert.deepEqual(shapeAt(doc, 1), ['i:mark']);
  run(h, insertText(doc, pos(2, 0), 'none'));
  assert.deepEqual(shapeAt(doc, 2), [':none']);
});

test('a pending format is laid over the inherited one and makes a run of its own', async () => {
  const doc = await open(p(r('abcd', '<w:i/>')));
  const h = history(doc);
  run(h, insertText(doc, pos(0, 2), 'X', { bold: true }));
  assert.deepEqual(shapeAt(doc, 0), ['i:ab', 'bi:X', 'i:cd']);
  // Typing on without a pending format continues the bold run to the left.
  run(h, insertText(doc, pos(0, 3), 'Y'));
  assert.deepEqual(shapeAt(doc, 0), ['i:ab', 'bi:XY', 'i:cd']);
  // Turning it off again joins the neighbours of equal formatting.
  run(h, insertText(doc, pos(0, 4), 'Z', { bold: null }));
  assert.deepEqual(shapeAt(doc, 0), ['i:ab', 'bi:XY', 'i:Zcd']);
});

test('tabs, line breaks and paragraph ends in typed text', async () => {
  const doc = await open(p(r('headtail', '<w:b/>'), '<w:jc w:val="center"/>' + SECT.replace('<w:pgSz', '<w:type w:val="continuous"/><w:pgSz')) + p('after'));
  const h = history(doc);
  const after = run(h, insertText(doc, pos(0, 4), 'a\tb\vc\nmiddle\r\n\nlast'));
  assert.deepEqual(texts(doc), ['heada\tb\nc', 'middle', '', 'lasttail', 'after']);
  assert.deepEqual(shapeAt(doc, 0), ['b:heada<tab>b<break>c']);
  assert.deepEqual(shapeAt(doc, 2), []);
  assert.deepEqual(after.focus, pos(3, 4));
  // All four are formatted alike; the section break stays at the end, with the last of them.
  for (let i = 0; i < 4; i++) assert.equal(paragraphAt(doc, i).props.align, 'center');
  assert.equal(paragraphAt(doc, 0).props.section, undefined);
  assert.equal(paragraphAt(doc, 2).props.section, undefined);
  assert.ok(paragraphAt(doc, 3).props.section);
  h.undo();
  assert.deepEqual(texts(doc), ['headtail', 'after']);
});

test('characters XML cannot hold are dropped', async () => {
  const doc = await open(p('ab'));
  run(history(doc), insertText(doc, pos(0, 1), 'x\u0000y\u0007\ud800z\udc00￿w'));
  assert.deepEqual(texts(doc), ['axyzwb']);
  // Text that is nothing but such characters changes nothing.
  assert.equal(insertText(doc, pos(0, 1), '\u0001').ops.length, 0);
});

test('typing over a range replaces its content', async () => {
  const doc = await open(p(r('one ', '<w:b/>') + r('two three')) + p('four five'));
  const h = history(doc);
  const after = run(h, insertText(doc, range(pos(1, 4), pos(0, 2)), 'X'));
  assert.deepEqual(texts(doc), ['onX five']);
  assert.deepEqual(shapeAt(doc, 0), ['b:onX', ': five']);
  assert.deepEqual(after.focus, pos(0, 3));
  assert.deepEqual(after.anchor, after.focus);
});

test('deleting within a paragraph joins what is left of equally formatted runs', async () => {
  const doc = await open(p(r('aaa') + r('BBB', '<w:b/>') + r('ccc')));
  const h = history(doc);
  const after = run(h, deleteRange(doc, range(pos(0, 2), pos(0, 7))));
  assert.deepEqual(shapeAt(doc, 0), [':aacc']);
  assert.deepEqual(after.focus, pos(0, 2));
  assert.equal(deleteRange(doc, range(pos(0, 1), pos(0, 1))).ops.length, 0);
});

test('deleting across paragraphs joins them: formatting of the first, unless that is empty; the section break of the last', async () => {
  const body =
    p('first', '<w:jc w:val="center"/>') + p('gone') + p('last one', '<w:jc w:val="right"/>' + SECT.replace('<w:pgSz', '<w:type w:val="continuous"/><w:pgSz')) + p('', '<w:jc w:val="center"/>') + p('tail', '<w:jc w:val="right"/>');
  const doc = await open(body);
  const h = history(doc);
  run(h, deleteRange(doc, range(pos(0, 3), pos(2, 5))));
  assert.deepEqual(texts(doc), ['firone', '', 'tail']);
  assert.equal(paragraphAt(doc, 0).props.align, 'center');
  assert.ok(paragraphAt(doc, 0).props.section, 'the section break of the last paragraph stays');
  // An empty first paragraph takes the formatting of the paragraph that joins it.
  run(h, deleteRange(doc, range(pos(1, 0), pos(2, 1))));
  assert.deepEqual(texts(doc), ['firone', 'ail']);
  assert.equal(paragraphAt(doc, 1).props.align, 'right');
});

const TABLE = tbl(tr(tc('a1') + tc('b1') + tc('c1')) + tr(tc('a2') + tc('b2') + tc('c2')), [1000, 1000, 1000]);

test('a range between cells empties the cells and keeps the table', async () => {
  const doc = await open(p('before') + TABLE + p('after'));
  const h = history(doc);
  const after = run(h, deleteRange(doc, range(pos([1, 0, 1, 0], 1), pos([1, 1, 1, 0], 1))));
  assert.deepEqual(texts(doc), ['before', 'a1', '', 'c1', 'a2', '', 'c2', 'after']);
  assert.deepEqual(after.focus, pos([1, 0, 1, 0], 0));
  assert.equal((doc.body.blocks[1] as Table).rows.length, 2);
  // Within one row.
  run(h, deleteRange(doc, range(pos([1, 1, 0, 0], 0), pos([1, 1, 2, 0], 1))));
  assert.deepEqual(texts(doc), ['before', 'a1', '', 'c1', '', '', '', 'after']);
  h.undo();
  h.undo();
  assert.deepEqual(texts(doc), ['before', 'a1', 'b1', 'c1', 'a2', 'b2', 'c2', 'after']);
});

test('a range that ends inside a table empties the cells it reaches; one that spans the table removes it', async () => {
  const doc = await open(p('before') + TABLE + p('after') + TABLE + p('end'));
  const h = history(doc);
  run(h, deleteRange(doc, range(pos(0, 3), pos([1, 0, 1, 0], 1))));
  assert.deepEqual(texts(doc), ['bef', '', '', 'c1', 'a2', 'b2', 'c2', 'after', 'a1', 'b1', 'c1', 'a2', 'b2', 'c2', 'end']);
  run(h, deleteRange(doc, range(pos([1, 1, 1, 0], 0), pos(2, 2))));
  assert.deepEqual(texts(doc), ['bef', '', '', 'c1', 'a2', '', '', 'ter', 'a1', 'b1', 'c1', 'a2', 'b2', 'c2', 'end']);
  // From before the second table to after it: the table goes, the paragraphs join.
  const after = run(h, deleteRange(doc, range(pos(2, 1), pos(4, 2))));
  assert.deepEqual(texts(doc), ['bef', '', '', 'c1', 'a2', '', '', 'td']);
  assert.deepEqual(after.focus, pos(2, 1));
  assert.equal(doc.body.blocks.filter((block) => block.type === 'table').length, 1);
});

test('two tables never end up touching', async () => {
  const doc = await open(p('x') + TABLE + p('between') + p('more') + TABLE + p('end'));
  run(history(doc), deleteRange(doc, range(pos([1, 1, 2, 0], 1), pos([4, 0, 0, 0], 1))));
  assert.deepEqual(doc.body.blocks.map((block) => block.type), ['paragraph', 'table', 'paragraph', 'table', 'paragraph']);
  assert.deepEqual(texts(doc), ['x', 'a1', 'b1', 'c1', 'a2', 'b2', '', '', '', 'b1', 'c1', 'a2', 'b2', 'c2', 'end']);
});

const DRAWING = '<w:r><w:drawing><wp:inline><wp:extent cx="95250" cy="95250"/><wp:docPr id="1" name="d"/></wp:inline></w:drawing></w:r>';
const field = (code: string, result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r(result)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

test('a drawing and a note reference are deleted as units of their own', async () => {
  const doc = await open({
    body: p(r('ab') + DRAWING + r('cd') + '<w:r><w:footnoteReference w:id="2"/></w:r>'),
    footnotes: '<w:footnote w:id="2"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:footnote>',
  });
  const h = history(doc);
  assert.equal(paragraphText(paragraphAt(doc, 0)), 'abcd');
  // The caret at 2 stands after the drawing: Backspace takes the drawing, not the "b".
  run(h, deleteBackward(doc, pos(0, 2)));
  assert.deepEqual(shapeAt(doc, 0), [':abcd', ':<footnoteReference>']);
  run(h, deleteBackward(doc, pos(0, 4)));
  assert.deepEqual(shapeAt(doc, 0), [':abcd']);
  run(h, deleteBackward(doc, pos(0, 4)));
  assert.deepEqual(shapeAt(doc, 0), [':abc']);
  h.undo();
  h.undo();
  h.undo();
  // Delete at 1 takes the "b" and leaves the drawing that follows it.
  run(h, deleteForward(doc, pos(0, 1)));
  assert.deepEqual(shapeAt(doc, 0), [':a', ':<drawing>', ':cd', ':<footnoteReference>']);
  // A range takes the drawings inside it.
  run(h, deleteRange(doc, range(pos(0, 0), pos(0, 2))));
  assert.deepEqual(shapeAt(doc, 0), [':d', ':<footnoteReference>']);
});

test('a computed field is one unit; another field goes when its result does, and stays when only part of it goes', async () => {
  const doc = await open(p(r('page ') + field('PAGE', '12') + r(' of ') + field('REF mark \\h', 'target') + r('.')));
  const h = history(doc);
  assert.equal(paragraphText(paragraphAt(doc, 0)), 'page 12 of target.');
  // Backspace after the page number takes the whole field, marks and code included.
  run(h, deleteBackward(doc, pos(0, 7)));
  assert.deepEqual(shapeAt(doc, 0).slice(0, 2), [':page  of ', ':<begin>']);
  h.undo();
  // A range that ends inside the result of a computed field takes all of it.
  run(h, deleteRange(doc, range(pos(0, 4), pos(0, 6))));
  assert.equal(paragraphText(paragraphAt(doc, 0)), 'page of target.');
  h.undo();
  // Text typed "inside" the page number goes after the field.
  run(h, insertText(doc, pos(0, 6), 'X'));
  assert.equal(paragraphText(paragraphAt(doc, 0)), 'page 12X of target.');
  h.undo();
  // Part of the result of an ordinary field: the field stays, with what is left of its result.
  run(h, deleteRange(doc, range(pos(0, 11), pos(0, 14))));
  assert.deepEqual(shapeAt(doc, 0).slice(6), [': of ', ':<begin>', ':<instrText>', ':<separate>', ':get', ':<end>', ':.']);
  // Backspace through the rest of the result: with its last character the field goes.
  run(h, deleteBackward(doc, pos(0, 14)));
  run(h, deleteBackward(doc, pos(0, 13)));
  run(h, deleteBackward(doc, pos(0, 12)));
  assert.deepEqual(shapeAt(doc, 0).slice(6), [': of .']);
});

test('typing at the edges of a field or a hyperlink stays outside, inside it stays inside', async () => {
  const doc = await open({
    body: p(r('a') + `<w:hyperlink r:id="rId9">${r('link')}</w:hyperlink>` + r('b') + field('REF x', 'ref') + r('c')),
    rels: [['rId9', 'hyperlink', 'https://example.com/', true]],
  });
  const h = history(doc);
  run(h, insertText(doc, pos(0, 1), '1'));
  run(h, insertText(doc, pos(0, 6), '2'));
  run(h, insertText(doc, pos(0, 4), '3'));
  assert.deepEqual(shapeAt(doc, 0).slice(0, 3), [':a1', ':li3nk@hyperlink', ':2b']);
  run(h, insertText(doc, pos(0, 9), '4'));
  run(h, insertText(doc, pos(0, 13), '5'));
  run(h, insertText(doc, pos(0, 11), '6'));
  assert.deepEqual(shapeAt(doc, 0).slice(2), [':2b4', ':<begin>', ':<instrText>', ':<separate>', ':r6ef', ':<end>', ':5c']);
});

test('a hyperlink left without text goes; bookmarks, comment marks and kept XML stay where they were', async () => {
  const doc = await open({
    body: p(
      r('a') +
        `<w:hyperlink r:id="rId9"><w:bookmarkStart w:id="1" w:name="in"/>${r('link')}<w:bookmarkEnd w:id="1"/></w:hyperlink>` +
        '<w:commentRangeStart w:id="3"/>' +
        r('b') +
        '<w:strange w:val="1"/><w:r><w:t>c</w:t><w:mystery/></w:r><w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r>' +
        r('d'),
    ),
    rels: [['rId9', 'hyperlink', 'https://example.com/', true]],
  });
  const h = history(doc);
  assert.deepEqual(shapeAt(doc, 0), [':a', '<bookmarkStart>@hyperlink', ':link@hyperlink', '<bookmarkEnd>@hyperlink', '<commentRangeStart>', ':b', '<unknown>', ':c<unknown>', '<commentRangeEnd>', ':<commentReference>', ':d']);
  run(h, deleteRange(doc, range(pos(0, 0), pos(0, 8))));
  assert.deepEqual(shapeAt(doc, 0), ['<bookmarkStart>', '<bookmarkEnd>', '<commentRangeStart>', '<unknown>', ':<unknown>', '<commentRangeEnd>', ':<commentReference>']);
  h.undo();
  // Part of the link: the link stays around what is left.
  run(h, deleteRange(doc, range(pos(0, 2), pos(0, 4))));
  assert.deepEqual(shapeAt(doc, 0).slice(0, 4), [':a', '<bookmarkStart>@hyperlink', ':lk@hyperlink', '<bookmarkEnd>@hyperlink']);
});

test('bookmarks of removed paragraphs move into the paragraph that stays', async () => {
  const doc = await open(p('one') + p('<w:bookmarkStart w:id="1" w:name="x"/>' + r('two')) + '<w:bookmarkStart w:id="2" w:name="y"/>' + p(r('three') + '<w:bookmarkEnd w:id="1"/>') + p('four'));
  run(history(doc), deleteRange(doc, range(pos(0, 1), pos(3, 4))));
  assert.deepEqual(shapeAt(doc, 0), [':o', '<bookmarkStart>', ':e', '<bookmarkEnd>']);
  // The mark between the blocks follows the joined paragraph.
  assert.deepEqual(doc.body.blocks.map((block) => block.type), ['paragraph', 'mark', 'paragraph']);
});

test('tracked deletions are not part of the text and survive edits around them', async () => {
  const doc = await open(p(r('ab') + '<w:del w:id="1" w:author="A"><w:r><w:delText>GONE</w:delText></w:r></w:del><w:ins w:id="2" w:author="A"><w:r><w:t>new</w:t></w:r></w:ins>' + r('cd')));
  const h = history(doc);
  assert.equal(paragraphText(paragraphAt(doc, 0)), 'abnewcd');
  run(h, deleteRange(doc, range(pos(0, 1), pos(0, 6))));
  assert.deepEqual(shapeAt(doc, 0), [':a', ':<deletedText>@del', ':d']);
  h.undo();
  // Inside the insertion the text stays in it; at its edge it does not.
  run(h, insertText(doc, pos(0, 3), 'X'));
  run(h, insertText(doc, pos(0, 6), 'Y'));
  assert.deepEqual(shapeAt(doc, 0), [':ab', ':<deletedText>@del', ':nXew@ins', ':Ycd']);
});

test('Backspace at the start of a paragraph: out of the list, then the indent, then joining', async () => {
  const doc = await open(p('zero') + p('item', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="720" w:firstLine="360"/>') + p('styled', '<w:pStyle w:val="ListStyle"/>'));
  const h = history(doc);
  run(h, deleteBackward(doc, pos(1, 0)));
  assert.equal(paragraphAt(doc, 1).props.numbering, undefined);
  assert.deepEqual(texts(doc), ['zero', 'item', 'styled']);
  run(h, deleteBackward(doc, pos(1, 0)));
  assert.deepEqual(paragraphAt(doc, 1).props.indent, { left: 720 });
  run(h, deleteBackward(doc, pos(1, 0)));
  assert.equal(paragraphAt(doc, 1).props.indent, undefined);
  const after = run(h, deleteBackward(doc, pos(1, 0)));
  assert.deepEqual(texts(doc), ['zeroitem', 'styled']);
  assert.deepEqual(after.focus, pos(0, 4));
  // A list that comes from the style is switched off on the paragraph.
  run(h, deleteBackward(doc, pos(1, 0)));
  assert.deepEqual(paragraphAt(doc, 1).props.numbering, { numId: 0 });
  assert.equal(paragraphAt(doc, 1).props.style, 'ListStyle');
});

test('Backspace and Delete do not cross the edges of cells and tables', async () => {
  const doc = await open(p('before') + TABLE + p('after'));
  assert.equal(deleteBackward(doc, pos([1, 0, 1, 0], 0)).ops.length, 0);
  assert.equal(deleteBackward(doc, pos(2, 0)).ops.length, 0);
  assert.equal(deleteForward(doc, pos(0, 6)).ops.length, 0);
  assert.equal(deleteForward(doc, pos([1, 1, 2, 0], 2)).ops.length, 0);
  assert.equal(deleteBackward(doc, pos(0, 0)).ops.length, 0);
  assert.equal(deleteForward(doc, pos(2, 5)).ops.length, 0);
});

test('Delete at the end of a paragraph joins the next one and leaves its drawings alone', async () => {
  const doc = await open(p('one') + p(DRAWING + r('two')));
  const h = history(doc);
  const after = run(h, deleteForward(doc, pos(0, 3)));
  assert.deepEqual(shapeAt(doc, 0), [':one', ':<drawing>', ':two']);
  assert.deepEqual(after.focus, pos(0, 3));
});

test('characters are deleted whole: surrogate pairs and letters with their marks; words by Ctrl', async () => {
  const doc = await open(p('a😀é̂z one  two, three'));
  const h = history(doc);
  run(h, deleteBackward(doc, pos(0, 6)));
  assert.equal(texts(doc)[0], 'a😀z one  two, three');
  run(h, deleteBackward(doc, pos(0, 3)));
  assert.equal(texts(doc)[0], 'az one  two, three');
  // A position inside a pair is taken to be after it.
  h.undo();
  run(h, insertText(doc, pos(0, 2), '!'));
  assert.equal(texts(doc)[0], 'a😀!z one  two, three');
  h.undo();
  h.undo();
  h.undo();
  run(h, deleteForward(doc, pos(0, 1)));
  run(h, deleteForward(doc, pos(0, 1)));
  assert.equal(texts(doc)[0], 'az one  two, three');
  run(h, deleteBackward(doc, pos(0, 11), 'word'));
  assert.equal(texts(doc)[0], 'az one  , three');
  run(h, deleteForward(doc, pos(0, 3), 'word'));
  assert.equal(texts(doc)[0], 'az , three');
});

test('Enter: both halves formatted alike, the section break with the second', async () => {
  const doc = await open(p(r('head', '<w:b/>') + r('tail'), '<w:jc w:val="center"/><w:rPr><w:i/></w:rPr>' + SECT.replace('<w:pgSz', '<w:type w:val="continuous"/><w:pgSz')));
  const h = history(doc);
  const after = run(h, splitParagraph(doc, pos(0, 4)));
  assert.deepEqual(texts(doc), ['head', 'tail']);
  assert.deepEqual(after.focus, pos(1, 0));
  assert.equal(paragraphAt(doc, 0).props.align, 'center');
  assert.equal(paragraphAt(doc, 1).props.align, 'center');
  assert.equal(paragraphAt(doc, 0).props.section, undefined);
  assert.ok(paragraphAt(doc, 1).props.section);
  assert.ok(paragraphAt(doc, 0).props.markProps && paragraphAt(doc, 0).props.markProps!.italic);
});

test('Enter at the end of a heading starts a paragraph of the following style; in the middle it does not', async () => {
  const doc = await open(p('Title', '<w:pStyle w:val="Heading1"/><w:jc w:val="center"/>') + p('quote', '<w:pStyle w:val="Quote"/>'));
  const h = history(doc);
  run(h, splitParagraph(doc, pos(0, 5)));
  assert.equal(paragraphAt(doc, 0).props.style, 'Heading1');
  assert.deepEqual(paragraphAt(doc, 1).props, {});
  run(h, splitParagraph(doc, pos(0, 2)));
  assert.deepEqual(texts(doc), ['Ti', 'tle', '', 'quote']);
  assert.equal(paragraphAt(doc, 1).props.style, 'Heading1');
  run(h, splitParagraph(doc, pos(3, 5)));
  assert.equal(paragraphAt(doc, 4).props.style, 'Heading1');
});

test('Enter in an empty list item takes it out of the list', async () => {
  const doc = await open(p('item', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') + p('', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'));
  const h = history(doc);
  const after = run(h, splitParagraph(doc, pos(1, 0)));
  assert.equal(doc.body.blocks.filter((block) => block.type === 'paragraph').length, 2);
  assert.equal(paragraphAt(doc, 1).props.numbering, undefined);
  assert.deepEqual(after.focus, pos(1, 0));
  // In an item with text the list goes on.
  run(h, splitParagraph(doc, pos(0, 4)));
  assert.deepEqual(paragraphAt(doc, 1).props.numbering, { numId: 1, level: 0 });
});

test('a page break ends the paragraph it is put in', async () => {
  const doc = await open(p('onetwo'));
  const after = run(history(doc), insertPageBreak(doc, pos(0, 3)));
  assert.deepEqual(shapeAt(doc, 0), [':one<break>']);
  assert.deepEqual(texts(doc), ['one\n', 'two']);
  assert.deepEqual(after.focus, pos(1, 0));
});

test('insertBlock puts a block before, after or in the middle of a paragraph', async () => {
  const doc = await open(p('onetwo'));
  const h = history(doc);
  const block = (text: string): Paragraph => ({ type: 'paragraph', props: {}, children: [{ type: 'run', props: {}, content: [{ type: 'text', text }] }] });
  run(h, insertBlock(doc, pos(0, 3), block('mid')));
  run(h, insertBlock(doc, pos(0, 0), block('top')));
  const after = run(h, insertBlock(doc, pos(3, 3), block('end')));
  assert.deepEqual(texts(doc), ['top', 'one', 'mid', 'two', 'end']);
  assert.deepEqual(after.focus, pos(4, 0));
});

test('blocks of a content control lose the source span of the control together', async () => {
  const doc = await open(p('zero') + `<w:sdt><w:sdtPr><w:alias w:val="c"/></w:sdtPr><w:sdtContent>${p('one')}${p('two')}</w:sdtContent></w:sdt>` + p('three'));
  const h = history(doc);
  const [, one, two] = doc.body.blocks;
  assert.ok(one.wrap && one.wrap === two.wrap && one.wrap.srcStart !== undefined && one.srcStart === undefined);
  run(h, insertText(doc, pos(1, 0), 'X'));
  const [zero, a, b, three] = doc.body.blocks;
  assert.ok(a.wrap && a.wrap === b.wrap, 'the blocks still share one wrapper');
  assert.equal(a.wrap!.srcStart, undefined);
  assert.equal(a.wrap!.kind, 'sdt');
  assert.ok(zero.srcStart !== undefined && three.srcStart !== undefined);
  assert.deepEqual(texts(doc), ['zero', 'Xone', 'two', 'three']);
  // Enter inside the control: the new paragraph is in it too.
  run(h, splitParagraph(doc, pos(1, 2)));
  assert.ok(doc.body.blocks.slice(1, 4).every((block) => block.wrap === doc.body.blocks[1].wrap));
  h.undo();
  h.undo();
  assert.equal(doc.body.blocks[1], one);
  assert.equal(doc.body.blocks[2], two);
});

test('notes and headers are stories of their own', async () => {
  const doc = await open({
    body: p(r('text') + '<w:r><w:footnoteReference w:id="2"/></w:r>') + '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="9000" w:h="9000"/></w:sectPr>',
    footnotes: '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> note</w:t></w:r></w:p></w:footnote>',
    headers: { rIdH: p('header') },
  });
  const h = history(doc);
  const note = doc.footnotes!.items.get(2)!;
  // Offset 0 of a note is after its number.
  run(h, insertText(doc, pos(0, 0, 'footnote:2'), 'A'));
  assert.deepEqual(shapeAt(doc, 0, 'footnote:2'), [':<footnoteRef>', ':A note']);
  assert.notEqual(doc.footnotes!.items.get(2), note);
  assert.equal(doc.footnotes!.items.get(2)!.srcStart, undefined);
  assert.equal(doc.footnotes!.modified, true);
  assert.equal(doc.body.modified, undefined);
  run(h, splitParagraph(doc, pos(0, 3, 'header:rIdH')));
  assert.deepEqual(texts(doc, 'header:rIdH'), ['hea', 'der']);
  assert.equal(doc.headers.get('rIdH')!.modified, true);
  h.undo();
  h.undo();
  assert.equal(doc.footnotes!.items.get(2), note);
  assert.equal(doc.footnotes!.modified, undefined);
  assert.equal(doc.headers.get('rIdH')!.modified, undefined);
});

test('what cannot be done is refused with a RangeError and changes nothing', async () => {
  const box =
    '<w:r><w:drawing><wp:inline><wp:extent cx="190500" cy="190500"/><wp:docPr id="7" name="s"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx><w:txbxContent>' +
    p('boxed') +
    '</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  const doc = await open({ body: p(r('text') + box) + p('more'), headers: { rIdH: p('header') } });
  const before = docPicture(doc);
  const refused = (make: () => unknown) => {
    assert.throws(make, RangeError);
    assert.equal(docPicture(doc), before);
  };
  refused(() => insertText(doc, pos(0, 5), 'x'));
  refused(() => insertText(doc, pos(0, -1), 'x'));
  refused(() => insertText(doc, pos(0, 1.5), 'x'));
  refused(() => insertText(doc, pos(9, 0), 'x'));
  refused(() => insertText(doc, pos([], 0), 'x'));
  refused(() => insertText(doc, pos([0, 0, 0, 0], 0), 'x'));
  refused(() => insertText(doc, pos(0, 0, 'nowhere'), 'x'));
  refused(() => insertText(doc, pos(0, 0, 'footnote:5'), 'x'));
  refused(() => deleteRange(doc, range(pos(0, 0), pos(0, 0, 'header:rIdH'))));
  refused(() => insertText(doc, pos(0, 0), 5 as unknown as string));
  // The text of a text box can be addressed, and read, but not edited.
  assert.equal(paragraphText(paragraphAt(doc, [0, -1, 0, 0, 0])), 'boxed');
  refused(() => insertText(doc, pos([0, -1, 0, 0, 0], 1), 'x'));
  refused(() => deleteBackward(doc, pos([0, -1, 0, 0, 0], 1)));
  refused(() => splitParagraph(doc, pos([0, -1, 0, 0, 0], 1)));
  refused(() => deleteRange(doc, range(pos(0, 0), pos([0, -1, 0, 0, 0], 1))));
  // A command that fails half way leaves nothing behind: the first of two steps is taken back.
  refused(() => insertText(doc, range(pos(0, 1), pos(1, 2)), 'x'.repeat(3) + '\n'.repeat(1000001)));
});

test('a field over several paragraphs keeps its marks unless all of it goes', async () => {
  const body =
    p(r('a') + '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' + r('r1')) +
    p('mid') +
    p(r('r2') + '<w:r><w:fldChar w:fldCharType="end"/></w:r>' + r('b'));
  const doc = await open(body);
  const h = history(doc);
  // From inside the first line of the result to inside the last: the three paragraphs join, the field stays whole.
  run(h, deleteRange(doc, range(pos(0, 2), pos(2, 1))));
  assert.deepEqual(shapeAt(doc, 0), [':a', ':<begin>', ':<instrText>', ':<separate>', ':r2', ':<end>', ':b']);
  assert.deepEqual((paragraphAt(doc, 0).children[2] as { fieldCode?: unknown }).fieldCode, true);
  h.undo();
  // All of the result: the field goes with it.
  run(h, deleteRange(doc, range(pos(0, 1), pos(2, 2))));
  assert.deepEqual(shapeAt(doc, 0), [':ab']);
  h.undo();
  // A paragraph in the middle of the result can be emptied and joined without harm to the field.
  run(h, deleteRange(doc, range(pos(0, 3), pos(1, 3))));
  assert.deepEqual(texts(doc), ['ar1', 'r2b']);
  assert.deepEqual(shapeAt(doc, 1), [':r2', ':<end>', ':b']);
});

test('a simple field the engine computes is one unit too', async () => {
  const doc = await open(p(r('p. ') + '<w:fldSimple w:instr=" PAGE ">' + r('1') + r('2', '<w:b/>') + '</w:fldSimple>' + r(' end')));
  const h = history(doc);
  assert.deepEqual(shapeAt(doc, 0), [':p. ', ':1@simpleField', 'b:2@simpleField', ': end']);
  run(h, deleteBackward(doc, pos(0, 5)));
  assert.deepEqual(shapeAt(doc, 0), [':p.  end']);
  h.undo();
  run(h, deleteForward(doc, pos(0, 3)));
  assert.deepEqual(shapeAt(doc, 0), [':p.  end']);
  h.undo();
  run(h, insertText(doc, pos(0, 4), 'X'));
  // After the field, with the formatting of the text before the field: what is in a field does not hand its formatting on.
  assert.deepEqual(shapeAt(doc, 0), [':p. ', ':1@simpleField', 'b:2@simpleField', ':X end']);
});

test('content of mc:AlternateContent is written back whole, so it is not edited', async () => {
  const alternate = (inner: string, fallback: string) => `<mc:AlternateContent><mc:Choice Requires="wps">${inner}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
  const doc = await open(p(r('ab') + alternate(r('choice'), r('fb')) + r('cd')) + alternate(p('block choice'), p('block fallback')) + p('end'));
  const h = history(doc);
  const inlineText = paragraphText(paragraphAt(doc, 0));
  const kept = inlineText.slice(2, inlineText.length - 2);
  assert.ok(kept.length > 0);
  // A range over the element takes the text around it and leaves the element.
  run(h, deleteRange(doc, range(pos(0, 1), pos(0, inlineText.length - 1))));
  assert.equal(paragraphText(paragraphAt(doc, 0)), 'a' + kept + 'd');
  // Text typed at an offset inside it goes after it.
  run(h, insertText(doc, pos(0, 2), 'X'));
  assert.equal(paragraphText(paragraphAt(doc, 0)), 'a' + kept + 'Xd');
  // A paragraph inside such an element at block level refuses edits.
  const before = docPicture(doc);
  assert.equal(doc.body.blocks[1].wrap!.kind, 'alternate');
  assert.throws(() => insertText(doc, pos(1, 0), 'x'), RangeError);
  assert.throws(() => deleteRange(doc, range(pos(0, 0), pos(2, 1))), RangeError);
  assert.throws(() => splitParagraph(doc, pos(1, 2)), RangeError);
  assert.equal(docPicture(doc), before);
});
