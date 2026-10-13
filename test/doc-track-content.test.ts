// Pictures, links, blocks and pasted content in a document that tracks its
// changes: each is a tracked insertion (a link over text that is there, a
// tracked deletion of the text and an insertion of the linked text), which
// accepting, rejecting, undoing and saving treat like any other.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { writeDocx } from '../src/formats/docx/write';
import { EMPTY_PARA_PROPS, type DocModel, type Paragraph, type Table } from '../src/model/doc/types';
import {
  TRACK_UNSUPPORTED,
  acceptAll,
  deleteBackward,
  deleteRange,
  fragmentOf,
  insertBlock,
  insertFragment,
  insertHyperlink,
  insertImage,
  insertPageBreak,
  insertText,
  rejectAll,
  rejectRevision,
  revisionsOf,
  type DocEdit,
  type DocPosition,
} from '../src/ops/doc/index';
import { checkInvariants, docPicture, p, paragraphAt, pos, r, range, run, tbl, tc, texts, tr } from './doc-ops-fixtures';
import { ME, body, checkFile, checkReopened, open, shape } from './doc-track-fixtures';

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4]);
const IMAGE = { bytes: PNG, type: 'image/png', width: 95250, height: 95250 };
const cell = (table: number, row: number, column: number, offset = 0): DocPosition => ({ story: 'body', path: [table, row, column, 0], offset });

const refused = (make: () => DocEdit): void => {
  assert.throws(make, (error: unknown) => error instanceof RangeError && error.message === TRACK_UNSUPPORTED);
};

const allResolved = (doc: DocModel, edit: { skipped: number }): void => {
  assert.equal(edit.skipped, 0, 'no change is passed over');
  void doc;
};

// ------------------------------------------------------------------- pictures

test('tracked: an inserted picture sits in a tracked insertion; rejected it goes, accepted it is a plain picture', async () => {
  const o = await open(p('abcd'));
  const before = shape(o.doc);
  const picture = docPicture(o.doc);
  run(o.history, insertImage(o.doc, pos(0, 2), IMAGE, ME));
  const drawing = paragraphAt(o.doc, 0).children.find((inline) => inline.type === 'run' && inline.content.some((item) => item.type === 'drawing'));
  assert.ok(drawing, 'the picture is in the paragraph');
  const around = drawing && drawing.wrap;
  assert.equal(around && around.kind === 'ins' ? around.author : undefined, 'Me', 'inside an insertion of the editor');
  const xml = await checkFile(o);
  assert.ok(/<w:ins w:id="1" w:author="Me"[^>]*><w:r><w:drawing>/.test(xml), xml);
  assert.deepEqual(revisionsOf(o.doc).map((revision) => revision.kind), ['insert']);
  await checkReopened(o, '', false);
  const edit = rejectAll(o.doc);
  allResolved(o.doc, edit);
  run(o.history, edit);
  assert.equal(shape(o.doc), before, 'rejected, the paragraph is what it was');
  assert.equal((await body(o)).indexOf('<w:drawing'), -1);
  o.history.undo();
  run(o.history, acceptAll(o.doc));
  const plain = await body(o);
  assert.ok(plain.indexOf('<w:drawing') > 0 && plain.indexOf('<w:ins ') < 0, plain);
  while (o.history.canUndo) o.history.undo();
  assert.equal(docPicture(o.doc), picture);
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);
});

test('tracked: a picture is deleted as a tracked change, also the one the editor just put in; its drawing stays in the file', async () => {
  const o = await open(p('abcd'));
  const at = run(o.history, insertImage(o.doc, pos(0, 2), IMAGE, ME));
  run(o.history, deleteBackward(o.doc, at.focus, 'char', ME));
  const xml = await checkFile(o);
  assert.ok(/<w:ins [^>]*><w:del [^>]*><w:r><w:drawing>/.test(xml), xml);
  await checkReopened(o, '', false);
});

test('tracked: a picture goes into a paragraph inside the insertion of somebody else by cutting that in two', async () => {
  const o = await open(p(`<w:ins w:id="7" w:author="You" w:date="2020-01-01T00:00:00Z">${r('abcd')}</w:ins>`));
  run(o.history, insertImage(o.doc, pos(0, 2), IMAGE, ME));
  const xml = await checkFile(o);
  assert.equal((xml.match(/<w:ins [^>]*w:author="You"/g) || []).length, 2, xml);
  assert.equal((xml.match(/<w:ins [^>]*w:author="Me"/g) || []).length, 1, xml);
});

// ---------------------------------------------------------------------- links

const LINK = 'https://example.org/';

test('tracked: a link over text is the text deleted and the same text inserted inside the link', async () => {
  const o = await open(p('one two three'));
  const before = shape(o.doc);
  const picture = docPicture(o.doc);
  const at = run(o.history, insertHyperlink(o.doc, range(pos(0, 4), pos(0, 7)), LINK, ME));
  assert.deepEqual(texts(o.doc), ['one two three']);
  assert.deepEqual([at.anchor.offset, at.focus.offset], [4, 7], 'the linked text is selected');
  const xml = await checkFile(o);
  assert.ok(/<w:del w:id="\d+"[^>]*><w:r><w:delText>two<\/w:delText><\/w:r><\/w:del><w:hyperlink r:id="rId\d+"><w:ins w:id="\d+"[^>]*><w:r>(<w:rPr>.*?<\/w:rPr>)?<w:t>two<\/w:t><\/w:r><\/w:ins><\/w:hyperlink>/.test(xml), xml);
  assert.deepEqual(revisionsOf(o.doc).map((revision) => `${revision.kind}:${revision.text}`), ['delete:two', 'insert:two']);
  await checkReopened(o);
  const edit = rejectAll(o.doc);
  allResolved(o.doc, edit);
  run(o.history, edit);
  assert.equal(shape(o.doc), before, 'rejected, the text is what it was, without a link');
  o.history.undo();
  run(o.history, acceptAll(o.doc));
  const accepted = await body(o);
  assert.ok(/<w:hyperlink r:id="rId\d+"><w:r>(<w:rPr>.*?<\/w:rPr>)?<w:t>two<\/w:t><\/w:r><\/w:hyperlink>/.test(accepted), accepted);
  assert.equal(accepted.indexOf('<w:del'), -1, accepted);
  while (o.history.canUndo) o.history.undo();
  assert.equal(docPicture(o.doc), picture);
});

test('tracked: a link at a caret is its address, inserted inside the link', async () => {
  const o = await open(p('ab'));
  const before = shape(o.doc);
  run(o.history, insertHyperlink(o.doc, range(pos(0, 1), pos(0, 1)), LINK, ME));
  assert.deepEqual(texts(o.doc), ['a' + LINK + 'b']);
  const xml = await checkFile(o);
  assert.ok(/<w:hyperlink r:id="rId\d+"><w:ins [^>]*><w:r>/.test(xml), xml);
  assert.equal(xml.indexOf('<w:del'), -1, xml);
  run(o.history, rejectAll(o.doc));
  assert.equal(shape(o.doc), before);
});

test('tracked: another address for a link is the linked text deleted in the old link and inserted in the new one', async () => {
  const o = await open(p(r('see ') + `<w:hyperlink r:id="rIdL">${r('site')}</w:hyperlink>` + r(' now')), { rels: [['rIdL', 'hyperlink', 'https://old.example/', true]] });
  const before = shape(o.doc);
  run(o.history, insertHyperlink(o.doc, range(pos(0, 4), pos(0, 8)), LINK, ME));
  assert.deepEqual(texts(o.doc), ['see site now']);
  const xml = await checkFile(o);
  assert.ok(/<w:hyperlink r:id="rIdL"><w:del [^>]*><w:r><w:delText>site<\/w:delText><\/w:r><\/w:del><\/w:hyperlink><w:hyperlink r:id="rId\d+"><w:ins [^>]*><w:r>/.test(xml), xml);
  await checkReopened(o);
  run(o.history, rejectAll(o.doc));
  assert.equal(shape(o.doc), before, 'rejected, the old link is back');
  o.history.undo();
  run(o.history, acceptAll(o.doc));
  const accepted = await body(o);
  assert.equal(accepted.indexOf('rIdL'), -1, accepted);
  assert.equal((accepted.match(/<w:hyperlink /g) || []).length, 1, accepted);
});

test('tracked: linking text the editor typed in this session leaves no deletion; the text of others is kept, deleted', async () => {
  const o = await open(p('ab'));
  run(o.history, insertText(o.doc, pos(0, 1), 'new', undefined, ME));
  run(o.history, insertHyperlink(o.doc, range(pos(0, 1), pos(0, 4)), LINK, ME));
  const xml = await checkFile(o);
  assert.equal(xml.indexOf('<w:del'), -1, xml);
  assert.deepEqual(texts(o.doc), ['anewb']);

  const theirs = await open(p(r('a') + `<w:ins w:id="7" w:author="You" w:date="2020-01-01T00:00:00Z">${r('new')}</w:ins>` + r('b')));
  run(theirs.history, insertHyperlink(theirs.doc, range(pos(0, 1), pos(0, 4)), LINK, ME));
  const kept = await checkFile(theirs);
  assert.ok(/<w:ins w:id="7" w:author="You"[^>]*><w:del [^>]*><w:r><w:delText>new<\/w:delText>/.test(kept), kept);
});

test('tracked: a link over a picture, a note reference or a field is refused, and nothing changes', async () => {
  const field = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>today</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const o = await open(p(r('a') + field + r('b')));
  const before = docPicture(o.doc);
  const relationships = o.doc.body.relationships;
  refused(() => insertHyperlink(o.doc, range(pos(0, 0), pos(0, 7)), LINK, ME));
  assert.equal(docPicture(o.doc), before);
  assert.ok(o.doc.body.relationships === relationships, 'no relationship is left behind');
});

// --------------------------------------------------------------------- blocks

test('tracked: a block put in whole is a tracked insertion: a paragraph with its mark, a table with its rows', async () => {
  const o = await open(p('abcd') + p('end'));
  const before = shape(o.doc);
  const paragraph: Paragraph = { type: 'paragraph', props: EMPTY_PARA_PROPS, children: [{ type: 'run', props: {}, content: [{ type: 'text', text: 'new' }] }] };
  for (const offset of [0, 2, 4]) {
    run(o.history, insertBlock(o.doc, pos(0, offset), paragraph, ME));
    const xml = await checkFile(o, `offset ${offset}`);
    assert.ok(/<w:p><w:pPr><w:rPr><w:ins [^>]*\/><\/w:rPr><\/w:pPr><w:ins [^>]*><w:r><w:t>new<\/w:t><\/w:r><\/w:ins><\/w:p>/.test(xml), xml);
    await checkReopened(o, `offset ${offset}`);
    const edit = rejectAll(o.doc);
    allResolved(o.doc, edit);
    run(o.history, edit);
    assert.equal(shape(o.doc), before, `offset ${offset}: rejected`);
    while (o.history.canUndo) o.history.undo();
  }
  const table: Table = { type: 'table', props: {}, grid: [1000], rows: [{ props: {}, cells: [{ props: {}, blocks: [{ type: 'paragraph', props: EMPTY_PARA_PROPS, children: [] }] }] }] };
  run(o.history, insertBlock(o.doc, pos(0, 2), table, ME));
  const xml = await checkFile(o);
  assert.ok(xml.indexOf('<w:trPr><w:ins ') > 0, xml);
  run(o.history, rejectAll(o.doc));
  assert.equal(shape(o.doc), before);
  // What is neither is refused where changes are tracked.
  refused(() => insertBlock(o.doc, pos(0, 2), { type: 'mark', mark: { type: 'bookmarkStart', id: '1', name: 'b' } }, ME));
});

test('tracked: a page break is an insertion with the mark it makes; rejected, the paragraph is one again', async () => {
  const o = await open(p('abcd'));
  const before = shape(o.doc);
  run(o.history, insertPageBreak(o.doc, pos(0, 2), ME));
  await checkFile(o);
  run(o.history, rejectAll(o.doc));
  assert.equal(shape(o.doc), before);
});

// ---------------------------------------------------------------------- paste

const RICH = p(r('plain ') + r('bold', '<w:b/>') + r(' text'), '<w:jc w:val="center"/>') + tbl(tr(tc('c1') + tc('c2')), [1200, 1200]) + p(r('see ') + `<w:hyperlink r:id="rIdL">${r('link')}</w:hyperlink>` + r(' end'));
const RELS = { rels: [['rIdL', 'hyperlink', 'https://example.org/', true]] as [string, string, string, boolean][] };

test('tracked: pasted text with its formatting is one insertion inside the paragraph', async () => {
  const source = await open(RICH, RELS);
  const fragment = fragmentOf(source.doc, range(pos(0, 3), pos(0, 12)));
  const o = await open(p('abcd'));
  const before = shape(o.doc);
  run(o.history, insertFragment(o.doc, pos(0, 2), fragment, ME));
  assert.deepEqual(texts(o.doc), ['abin bold tcd']);
  const xml = await checkFile(o);
  assert.ok(/<w:ins w:id="1"[^>]*><w:r><w:t[^>]*>in <\/w:t><\/w:r><w:r><w:rPr><w:b\/><\/w:rPr><w:t>bold<\/w:t><\/w:r><w:r><w:t[^>]*> t<\/w:t><\/w:r><\/w:ins>/.test(xml), xml);
  await checkReopened(o);
  const edit = rejectAll(o.doc);
  allResolved(o.doc, edit);
  run(o.history, edit);
  assert.equal(shape(o.doc), before);
});

test('tracked: pasted paragraphs, a table and a link are tracked insertions, marks and rows included; rejected, nothing is left', async () => {
  const source = await open(RICH, RELS);
  const fragment = fragmentOf(source.doc, range(pos(0, 0), pos(2, 12)));
  for (const offset of [0, 2, 4]) {
    const o = await open(p('abcd') + p('tail'));
    const before = shape(o.doc);
    const picture = docPicture(o.doc);
    run(o.history, insertFragment(o.doc, pos(0, offset), fragment, ME));
    checkInvariants(o.doc, false);
    const xml = await checkFile(o, `offset ${offset}`);
    // The table: every row inserted. The link: its text inserted inside it. The marks of the pasted paragraphs: inserted.
    assert.equal((xml.match(/<w:trPr><w:ins /g) || []).length, 1, xml);
    assert.ok(/<w:hyperlink r:id="rId\d+"><w:ins [^>]*><w:r>/.test(xml), xml);
    assert.ok(xml.indexOf('<w:t>bold</w:t>') > 0, xml);
    assert.equal(xml.indexOf('<w:del'), -1, xml);
    // Nothing of what was pasted stands outside an insertion.
    for (const inserted of ['plain ', 'bold', 'c1', 'c2', 'link']) {
      const index = xml.indexOf(`>${inserted}</w:t>`);
      assert.ok(index > 0, `${inserted}\n${xml}`);
      assert.ok(xml.lastIndexOf('<w:ins ', index) > xml.lastIndexOf('</w:ins>', index), `offset ${offset}: "${inserted}" is not in an insertion\n${xml}`);
    }
    await checkReopened(o, `offset ${offset}`);
    const edit = rejectAll(o.doc);
    allResolved(o.doc, edit);
    run(o.history, edit);
    assert.equal(shape(o.doc), before, `offset ${offset}: rejected`);
    assert.deepEqual(revisionsOf(o.doc), []);
    o.history.undo();
    const accept = acceptAll(o.doc);
    allResolved(o.doc, accept);
    run(o.history, accept);
    assert.deepEqual(revisionsOf(o.doc), []);
    const plain = await checkFile(o, `offset ${offset} accepted`);
    assert.equal(plain.indexOf('<w:ins '), -1, plain);
    assert.ok(plain.indexOf('>c1</w:t>') > 0 && plain.indexOf('<w:hyperlink ') > 0, plain);
    while (o.history.canUndo) o.history.undo();
    assert.equal(docPicture(o.doc), picture);
  }
});

test('tracked: a paste over a selection deletes the selection as a change and inserts after it', async () => {
  const source = await open(RICH, RELS);
  const fragment = fragmentOf(source.doc, range(pos(0, 0), pos(2, 12)));
  const o = await open(p('abcd') + p('tail'));
  const before = shape(o.doc);
  run(o.history, insertFragment(o.doc, range(pos(0, 1), pos(1, 2)), fragment, ME));
  const xml = await checkFile(o);
  assert.ok(xml.indexOf('<w:delText>bcd</w:delText>') > 0 && xml.indexOf('<w:delText>ta</w:delText>') > 0, xml);
  await checkReopened(o);
  const edit = rejectAll(o.doc);
  allResolved(o.doc, edit);
  run(o.history, edit);
  assert.equal(shape(o.doc), before);
});

test('tracked: a paste into a table cell, and a pasted picture', async () => {
  const source = await open(p('x'));
  run(source.history, insertImage(source.doc, pos(0, 1), IMAGE));
  const fragment = fragmentOf(source.doc, range(pos(0, 0), pos(0, 1)));
  const o = await open(p('top') + tbl(tr(tc('cell')), [2000]) + p('end'));
  const before = shape(o.doc);
  run(o.history, insertFragment(o.doc, cell(1, 0, 0, 2), fragment, ME));
  const xml = await checkFile(o);
  assert.ok(/<w:ins [^>]*>(<w:r>.*?<\/w:r>)*<w:r><w:drawing>/.test(xml), xml);
  await checkReopened(o, '', false);
  run(o.history, rejectAll(o.doc));
  assert.equal(shape(o.doc), before);
});

test('tracked: what is pasted carries no revision of the document it was copied from', async () => {
  const source = await open(p(r('a') + `<w:ins w:id="50" w:author="You" w:date="2020-01-01T00:00:00Z">${r('theirs')}</w:ins>` + r('b'), '<w:rPr><w:ins w:id="51" w:author="You" w:date="2020-01-01T00:00:00Z"/></w:rPr>') + p('next'));
  const fragment = fragmentOf(source.doc, range(pos(0, 0), pos(1, 2)));
  const o = await open(p('abcd'));
  run(o.history, insertFragment(o.doc, pos(0, 2), fragment, ME));
  const xml = await checkFile(o);
  assert.equal(xml.indexOf('w:author="You"'), -1, xml);
  assert.ok(revisionsOf(o.doc).every((revision) => revision.author === 'Me'));
});

test('tracked: a cut is a tracked deletion, tables and pictures in it too', async () => {
  const o = await open(RICH, RELS);
  const before = shape(o.doc);
  run(o.history, deleteRange(o.doc, range(pos(0, 3), pos(2, 6)), ME));
  const xml = await checkFile(o);
  assert.equal((xml.match(/<w:trPr><w:del /g) || []).length, 1, xml);
  assert.ok(xml.indexOf('<w:delText>li</w:delText>') > 0, xml);
  run(o.history, rejectAll(o.doc));
  assert.equal(shape(o.doc), before);
});

test('a document that does not track pastes, links and puts pictures in as before', async () => {
  const source = await open(RICH, RELS);
  const fragment = fragmentOf(source.doc, range(pos(0, 0), pos(2, 12)));
  const o = await open(p('abcd') + p('tail'), { settings: '' });
  run(o.history, insertFragment(o.doc, pos(0, 2), fragment));
  run(o.history, insertHyperlink(o.doc, range(pos(0, 0), pos(0, 2)), LINK));
  run(o.history, insertImage(o.doc, pos(0, 1), IMAGE));
  const xml = await body(o);
  assert.equal(xml.indexOf('<w:ins '), -1, xml);
  assert.equal(xml.indexOf('<w:del'), -1, xml);
  assert.deepEqual(revisionsOf(o.doc), []);
  while (o.history.canUndo) o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);
});

test('tracked: one change of a paste resolved by itself leaves the others', async () => {
  const source = await open(RICH, RELS);
  const fragment = fragmentOf(source.doc, range(pos(0, 0), pos(2, 12)));
  const o = await open(p(`<w:del w:id="70" w:author="You" w:date="2020-01-01T00:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del>` + r('abcd')) + p('tail'));
  run(o.history, insertFragment(o.doc, pos(0, 2), fragment, ME));
  const mine = revisionsOf(o.doc).filter((revision) => revision.author === 'Me');
  assert.ok(mine.length >= 1, 'the paste is listed');
  run(o.history, rejectRevision(o.doc, mine[0]));
  assert.ok(revisionsOf(o.doc).some((revision) => revision.author === 'You' && revision.kind === 'delete'), 'the deletion of the other author is still there');
  await checkFile(o);
});
