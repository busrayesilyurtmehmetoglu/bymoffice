// Fourth review of tracked changes, the interplay of editing without tracking and the changes a
// document already holds: an edit that is not tracked makes no change of the list and takes none
// away but those whose content it removes, and the paragraph mark an edit makes is nobody's
// revision.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import { type DocModel } from '../src/model/doc/types';
import {
  DocHistory,
  acceptAll,
  acceptRevision,
  deleteBackward,
  deleteForward,
  deleteRange,
  insertColumn,
  insertPageBreak,
  insertRow,
  insertTable,
  insertText,
  rejectAll,
  rejectRevision,
  revisionsOf,
  splitParagraph,
  storyIds,
  storyOf,
  type DocEditOptions,
  type DocRange,
} from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { parseXml } from '../src/xml';
import { NUMBERING, STYLES, picture, p, paragraphAt, pos, r, range, run, tbl, tc, texts, tr } from './doc-ops-fixtures';
import { buildDocx as buildSpec, type DocxSpec } from './docx-fixtures';

const decoder = new TextDecoder();
const ME: DocEditOptions = { author: 'Me', date: '2026-10-09T10:00:00Z' };
const TRACKED: DocEditOptions = { ...ME, trackChanges: true };
const SECT = '<w:sectPr><w:pgSz w:w="9000" w:h="9000"/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr>';
const WHO = (id: number, author = 'Ann') => `w:id="${id}" w:author="${author}" w:date="2020-01-01T00:00:00Z"`;
const INS_MARK = (id: number, author = 'Ann') => `<w:rPr><w:ins ${WHO(id, author)}/></w:rPr>`;
const DEL_MARK = (id: number, author = 'Ann') => `<w:rPr><w:del ${WHO(id, author)}/></w:rPr>`;

interface Opened {
  pkg: OpcPackage;
  doc: DocModel;
  history: DocHistory<DocRange | undefined>;
  bytes: Uint8Array;
}

async function open(body: string, extra: Partial<DocxSpec> = {}): Promise<Opened> {
  const bytes = await buildSpec({ body: body + SECT, styles: STYLES, ...extra });
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  return { pkg, doc, history: new DocHistory<DocRange | undefined>(doc), bytes };
}

async function written(o: Opened): Promise<string> {
  await writeDocx(o.pkg, o.doc);
  return decoder.decode(await o.pkg.readPart('/word/document.xml'));
}

async function body(o: Opened): Promise<string> {
  const text = await written(o);
  return text.slice(text.indexOf('<w:body>') + 8, text.lastIndexOf('<w:sectPr>'));
}

/** The document as a reader of the file written now has it. */
async function reread(o: Opened): Promise<DocModel> {
  await writeDocx(o.pkg, o.doc);
  return readDocx(await OpcPackage.open(await o.pkg.save()));
}

/** The picture of a document without what only says where a block stood in its source. */
function shapeOf(doc: DocModel): string {
  const blocks = new Map<string, unknown>();
  for (const id of storyIds(doc)) blocks.set(id, storyOf(doc, id).blocks);
  return picture({ blocks, section: doc.body.section, settings: doc.settings })
    .replace(/"src(Start|End)":\d+,?/g, '')
    // The second half of a control a paragraph break cut in two is written without the id of the control.
    .replace(/\{"attrs":\{"map":\[\["w:val","99"\]\]\},"children":\[\],"name":"w:id","type":"element"\}/g, '')
    .replace(/,}/g, '}');
}

/** The model is what a reader of the file written from it has. */
async function assertRereads(o: Opened, label = ''): Promise<void> {
  const a = shapeOf(o.doc);
  const b = shapeOf(await reread(o));
  if (a === b) return;
  let i = 0;
  while (i < a.length && a.charAt(i) === b.charAt(i)) i++;
  assert.fail(`${label}: the model is not what the file reads as, from ${i}\nmodel: ${a.slice(Math.max(0, i - 200), i + 200)}\nfile:  ${b.slice(Math.max(0, i - 200), i + 200)}`);
}

const count = (text: string, part: string): number => text.split(part).length - 1;

/** The ids of every element of a tracked change in a text of XML. */
const idsIn = (xml: string): string[] => Array.from(xml.matchAll(/<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|numberingChange|trPrChange|tcPrChange|tblPrChange|cellIns|cellDel)\b[^>]*? w:id="(\d+)"/g), (match) => match[1]);

function assertUniqueIds(xml: string): void {
  const ids = idsIn(xml);
  assert.equal(new Set(ids).size, ids.length, `an id twice: ${ids.join(',')}\n${xml}`);
}

// ------------------------------------------------------------------------- A1

test('review 4 (A1): Enter without tracking in a paragraph whose mark is a tracked insertion makes a plain mark', async () => {
  const o = await open(p(r('aaaa bbbb'), INS_MARK(9)) + p('tail'));
  assert.equal(revisionsOf(o.doc).length, 1);
  run(o.history, splitParagraph(o.doc, pos(0, 4)));
  assert.deepEqual(texts(o.doc), ['aaaa', ' bbbb', 'tail']);
  // The mark the file had is the end of the paragraph: it is with the second half. The mark the user made is plain.
  assert.equal(paragraphAt(o.doc, 0).props.markProps, undefined);
  assert.ok(paragraphAt(o.doc, 1).props.markProps?.inserted);
  const list = revisionsOf(o.doc);
  assert.equal(list.length, 1, JSON.stringify(list));
  const xml = await body(o);
  assert.equal(count(xml, '<w:ins '), 1, xml);
  assertUniqueIds(xml);
  // Rejecting everything takes back what Ann did, not the Enter of the user.
  run(o.history, rejectAll(o.doc));
  assert.deepEqual(texts(o.doc), ['aaaa', ' bbbbtail']);
});

test('review 4 (A1): the record of changed formatting stays with the paragraph that keeps the mark of the file', async () => {
  const pPr = `<w:jc w:val="center"/><w:rPr><w:ins ${WHO(9)}/><w:b/><w:rPrChange ${WHO(10)}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:pPrChange ${WHO(11)}><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange>`;
  const o = await open(p(r('aaaa bbbb'), pPr) + p('tail'));
  assert.equal(revisionsOf(o.doc).length, 3);
  run(o.history, splitParagraph(o.doc, pos(0, 4)));
  const first = paragraphAt(o.doc, 0).props;
  const second = paragraphAt(o.doc, 1).props;
  // Both halves are formatted as the paragraph was; only the second says what was changed.
  assert.equal(first.align, 'center');
  assert.equal(first.change, undefined);
  assert.equal(first.markProps?.bold, true);
  assert.equal(first.markProps?.change, undefined);
  assert.equal(first.markProps?.inserted, undefined);
  assert.ok(second.change && second.markProps?.change && second.markProps.inserted);
  assert.equal(revisionsOf(o.doc).length, 3);
  const xml = await body(o);
  assertUniqueIds(xml);
  assert.equal(count(xml, '<w:pPrChange '), 1, xml);
  assert.equal(count(xml, '<w:rPrChange '), 1, xml);
  parseXml(await written(o));
  o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);
});

test('review 4 (A1): a paragraph whose mark is a tracked deletion, split without tracking, is not joined again by "accept all"', async () => {
  const o = await open(p(r('aaaa bbbb'), DEL_MARK(9)) + p('tail'));
  run(o.history, splitParagraph(o.doc, pos(0, 4)));
  assert.equal(revisionsOf(o.doc).length, 1);
  run(o.history, acceptAll(o.doc));
  assert.deepEqual(texts(o.doc), ['aaaa', ' bbbbtail']);
});

test('review 4 (A1): a page break and a table put into such a paragraph without tracking leave one change too', async () => {
  const breaks = await open(p(r('aaaa bbbb'), INS_MARK(9)) + p('tail'));
  run(breaks.history, insertPageBreak(breaks.doc, pos(0, 4)));
  assert.equal(revisionsOf(breaks.doc).length, 1);
  assert.equal(paragraphAt(breaks.doc, 0).props.markProps, undefined);
  assertUniqueIds(await body(breaks));

  const table = await open(p(r('aaaa bbbb'), INS_MARK(9)) + p('tail'));
  run(table.history, insertTable(table.doc, pos(0, 4), 1, 1));
  assert.equal(revisionsOf(table.doc).length, 1, JSON.stringify(revisionsOf(table.doc)));
  assert.equal(paragraphAt(table.doc, 0).props.markProps, undefined);
  assertUniqueIds(await body(table));

  // Text of several lines typed (pasted) into it.
  const lines = await open(p(r('aaaa bbbb'), INS_MARK(9)) + p('tail'));
  run(lines.history, insertText(lines.doc, pos(0, 4), 'x\ny\nz'));
  assert.deepEqual(texts(lines.doc), ['aaaax', 'y', 'z bbbb', 'tail']);
  assert.equal(revisionsOf(lines.doc).length, 1);
  assert.ok(paragraphAt(lines.doc, 2).props.markProps?.inserted);
  assertUniqueIds(await body(lines));
});

test('review 4 (A1): Enter at the end of a heading keeps the tracked mark of the heading on the paragraph that follows', async () => {
  const o = await open(p(r('Title'), '<w:pStyle w:val="Heading1"/>' + INS_MARK(9)) + p('tail'));
  const next = o.doc.styles.styles.get('Heading1')?.next;
  run(o.history, splitParagraph(o.doc, pos(0, 5)));
  assert.equal(revisionsOf(o.doc).length, 1);
  assert.equal(paragraphAt(o.doc, 0).props.markProps, undefined);
  assert.ok(paragraphAt(o.doc, 1).props.markProps?.inserted, `next style: ${next}`);
});

// ------------------------------------------------------------------------- A2

test('review 4 (A2): joining without tracking takes the revision of the mark that goes away with it (Delete at the end)', async () => {
  const o = await open(p(r('first'), DEL_MARK(9)) + p('second') + p('third'));
  assert.equal(revisionsOf(o.doc).length, 1);
  run(o.history, deleteForward(o.doc, pos(0, 5)));
  assert.deepEqual(texts(o.doc), ['firstsecond', 'third']);
  // The mark Ann deleted is the mark the user removed for good: nothing is left to resolve.
  assert.deepEqual(revisionsOf(o.doc), []);
  assert.equal(paragraphAt(o.doc, 0).props.markProps, undefined);
  run(o.history, acceptAll(o.doc));
  assert.deepEqual(texts(o.doc), ['firstsecond', 'third']);
  assert.equal(count(await body(o), '<w:del '), 0);
});

test('review 4 (A2): Backspace at the start of a paragraph removes the inserted mark before it, and "reject all" joins nothing more', async () => {
  const o = await open(p(r('aaaa bbbb'), INS_MARK(9)) + p('second para') + p('third'));
  run(o.history, deleteBackward(o.doc, pos(1, 0)));
  assert.deepEqual(texts(o.doc), ['aaaa bbbbsecond para', 'third']);
  assert.deepEqual(revisionsOf(o.doc), []);
  run(o.history, rejectAll(o.doc));
  assert.deepEqual(texts(o.doc), ['aaaa bbbbsecond para', 'third']);
});

test('review 4 (A2): the revision of the mark that stays is kept by a join without tracking', async () => {
  const o = await open(p('first') + p(r('second'), `<w:jc w:val="center"/><w:rPr><w:ins ${WHO(9)}/><w:rPrChange ${WHO(10)}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr>`) + p('third'));
  assert.equal(revisionsOf(o.doc).length, 2);
  run(o.history, deleteBackward(o.doc, pos(1, 0)));
  assert.deepEqual(texts(o.doc), ['firstsecond', 'third']);
  const props = paragraphAt(o.doc, 0).props;
  // The formatting is the first paragraph's (the rule of the editor); the mark is, as a change, what the second's was.
  assert.equal(props.align, undefined);
  assert.equal(props.markProps?.inserted?.author, 'Ann');
  assert.ok(props.markProps?.change);
  const list = revisionsOf(o.doc);
  assert.deepEqual(
    list.map((revision) => revision.kind),
    ['insert', 'format'],
  );
  assertUniqueIds(await body(o));
  // Ann's mark can still be rejected: the paragraph joins the one after it.
  run(o.history, rejectRevision(o.doc, list[0]));
  assert.deepEqual(texts(o.doc), ['firstsecondthird']);
});

test('review 4 (A2): a range deleted without tracking over tracked marks leaves the mark of its last paragraph', async () => {
  const o = await open(p(r('one'), INS_MARK(9)) + p(r('two'), DEL_MARK(10, 'Bob')) + p(r('three'), INS_MARK(11, 'Cid')) + p('four'));
  assert.equal(revisionsOf(o.doc).length, 3);
  run(o.history, deleteRange(o.doc, range(pos(0, 1), pos(2, 2))));
  assert.deepEqual(texts(o.doc), ['oree', 'four']);
  const list = revisionsOf(o.doc);
  assert.equal(list.length, 1, JSON.stringify(list));
  assert.equal(list[0].author, 'Cid');
  o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);
});

// ------------------------------------------------------------------------- A4

const REVISED_ROW =
  `<w:tr><w:tblPrEx><w:jc w:val="center"/><w:tblPrExChange ${WHO(20)}><w:tblPrEx/></w:tblPrExChange></w:tblPrEx><w:trPr><w:cantSplit/><w:ins ${WHO(21)}/><w:trPrChange ${WHO(22)}><w:trPr/></w:trPrChange></w:trPr>` +
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:cellIns ${WHO(23)}/><w:tcPrChange ${WHO(24)}><w:tcPr/></w:tcPrChange></w:tcPr>` +
  `<w:p><w:pPr><w:jc w:val="right"/><w:rPr><w:del ${WHO(25)}/><w:b/><w:rPrChange ${WHO(26)}><w:rPr/></w:rPrChange></w:rPr><w:pPrChange ${WHO(27)}><w:pPr/></w:pPrChange></w:pPr>${r('a')}</w:p><w:p>${r('a2')}</w:p></w:tc>` +
  tc('b') +
  '</w:tr>';

test('review 4 (A4): a row put beside a row with tracked changes carries none of them', async () => {
  for (const where of ['above', 'below'] as const) {
    const o = await open(p('before') + tbl(REVISED_ROW, [2000, 2000]) + p('after'));
    const before = revisionsOf(o.doc);
    const idsBefore = idsIn(await body(o));
    run(o.history, insertRow(o.doc, pos([1, 0, 0, 0], 0), where));
    const table = o.doc.body.blocks[1];
    assert.ok(table.type === 'table' && table.rows.length === 2);
    const row = table.rows[where === 'above' ? 0 : 1];
    // Formatted like the row beside it, and no change of anybody's.
    assert.equal(row.props.cantSplit, true);
    assert.equal(row.props.inserted, undefined);
    assert.equal(row.props.change, undefined);
    assert.equal(row.tableProps?.change, undefined);
    assert.equal(row.cells[0].props.change, undefined);
    assert.ok(!(row.cells[0].props.extras || []).some((extra) => /cellIns/.test(extra.el.name)));
    const para = row.cells[0].blocks[0];
    assert.ok(para.type === 'paragraph');
    assert.equal(para.props.align, 'right');
    assert.equal(para.props.change, undefined);
    assert.deepEqual(para.props.markProps, { bold: true });
    const after = revisionsOf(o.doc);
    assert.equal(after.length, before.length, `${where}: ${JSON.stringify(after.map((revision) => revision.kind))}`);
    const xml = await body(o);
    assertUniqueIds(xml);
    assert.deepEqual(idsIn(xml).sort(), idsBefore.sort(), where);
    parseXml(await written(o));
    o.history.undo();
    await writeDocx(o.pkg, o.doc);
    assert.deepEqual(await o.pkg.save(), o.bytes, where);
  }
});

test('review 4 (A4): a column put beside cells with tracked changes carries none of them', async () => {
  for (const side of ['left', 'right'] as const) {
    const o = await open(p('before') + tbl(REVISED_ROW, [2000, 2000]) + p('after'));
    const before = revisionsOf(o.doc).length;
    const idsBefore = idsIn(await body(o));
    run(o.history, insertColumn(o.doc, pos([1, 0, 0, 0], 0), side));
    const table = o.doc.body.blocks[1];
    assert.ok(table.type === 'table' && table.rows[0].cells.length === 3);
    const cell = table.rows[0].cells[side === 'left' ? 0 : 1];
    assert.equal(cell.props.change, undefined);
    assert.ok(!(cell.props.extras || []).some((extra) => /cellIns/.test(extra.el.name)));
    const para = cell.blocks[0];
    assert.ok(para.type === 'paragraph');
    assert.equal(para.props.change, undefined);
    assert.deepEqual(para.props.markProps, { bold: true });
    // The list may group the changes of the row otherwise around a new plain cell: it never has fewer, and the file has no other ids.
    assert.ok(revisionsOf(o.doc).length >= before, side);
    const xml = await body(o);
    assertUniqueIds(xml);
    assert.deepEqual(idsIn(xml).sort(), idsBefore.sort(), side);
  }
});

test('review 4 (A4): with tracking on, rows, columns and tables are still refused', async () => {
  const o = await open(p('before') + tbl(REVISED_ROW, [2000, 2000]) + p('after'));
  assert.throws(() => insertRow(o.doc, pos([1, 0, 0, 0], 0), 'below', TRACKED), RangeError);
  assert.throws(() => insertColumn(o.doc, pos([1, 0, 0, 0], 0), 'right', TRACKED), RangeError);
  assert.throws(() => insertTable(o.doc, pos(0, 2), 1, 1, TRACKED), RangeError);
});

// ------------------------------------------------------------------------- A5

const LISTED = `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/><w:numberingChange ${WHO(30)} w:original="1."/><w:ins ${WHO(31)}/></w:numPr>`;

test('review 4 (A5): a change kept inside the properties stays with the half that keeps the mark, when the split is not tracked', async () => {
  const o = await open(p(r('aaaa bbbb'), LISTED) + p('tail'), { numbering: NUMBERING });
  const before = revisionsOf(o.doc);
  assert.equal(before.length, 1);
  assert.equal(before[0].kind, 'other');
  assert.equal(before[0].count, 2);
  run(o.history, splitParagraph(o.doc, pos(0, 4)));
  const xml = await body(o);
  assert.equal(count(xml, '<w:numberingChange '), 1, xml);
  assertUniqueIds(xml);
  // Both halves are items of the list; only the second says that the number was changed.
  assert.deepEqual(paragraphAt(o.doc, 0).props.numbering, { level: 0, numId: 2 });
  assert.deepEqual(paragraphAt(o.doc, 1).props.numbering, { level: 0, numId: 2 });
  assert.equal(count(xml, '<w:numId w:val="2"/>'), 2, xml);
  assert.ok(/<w:p><w:pPr><w:numPr><w:ilvl w:val="0"\/><w:numId w:val="2"\/><\/w:numPr><\/w:pPr>/.test(xml), xml);
  const after = revisionsOf(o.doc);
  assert.equal(after.length, 1);
  assert.equal(after[0].count, 2);
  // The file is one a second writing does not change.
  const again = await readDocx(await OpcPackage.open(await o.pkg.save()));
  assert.equal(revisionsOf(again).length, 1);
  assert.deepEqual(texts(again), texts(o.doc));
});

test('review 4 (A5): a tracked split gives the changes kept inside the properties of the new paragraph ids of their own', async () => {
  const o = await open(p(r('aaaa bbbb'), LISTED) + p('tail'), { numbering: NUMBERING });
  run(o.history, splitParagraph(o.doc, pos(0, 4), TRACKED));
  const xml = await body(o);
  assertUniqueIds(xml);
  parseXml(await written(o));
  await assertRereads(o);
  // Text of several lines, tracked: every paragraph it ends is one more.
  const lines = await open(p(r('aaaa bbbb'), LISTED) + p('tail'), { numbering: NUMBERING });
  run(lines.history, insertText(lines.doc, pos(0, 4), 'x\ny\nz', undefined, TRACKED));
  assertUniqueIds(await body(lines));
});

// ------------------------------------------------------------------------- A7

const OWN = (inner: string, id = 7) => `<w:ins w:id="${id}" w:author="Me" w:date="2020-01-01T00:00:00Z">${inner}</w:ins>`;
const CONTROL = (inner: string) => `<w:sdt><w:sdtPr><w:id w:val="99"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
const deletedText = (xml: string): string => Array.from(xml.matchAll(/<w:delText[^>]*>([^<]*)<\/w:delText>/g), (match) => match[1]).join('');

test('review 4 (A7): what is left of an own insertion that is all a control holds does not depend on how its runs are cut', async () => {
  const shapes = [r('abcd'), r('ab') + r('cd'), r('a') + r('b', '<w:b/>') + r('cd'), r('a') + r('b') + r('c') + r('d')];
  const results: string[] = [];
  for (const inner of shapes) {
    const o = await open(p(r('x') + CONTROL(OWN(inner)) + r('y')), { settings: '<w:trackRevisions/>' });
    run(o.history, deleteRange(o.doc, range(pos(0, 1), pos(0, 5)), ME));
    assert.deepEqual(texts(o.doc), ['xy']);
    const xml = await body(o);
    // The control would be left with nothing: all of the text stays, deleted, inside it.
    assert.equal(deletedText(xml), 'abcd', xml);
    assert.equal(count(xml, '<w:id w:val="99"/>'), 1, xml);
    assert.equal(count(xml, '<w:t'), 2, xml);
    assertUniqueIds(xml);
    results.push(deletedText(xml));
    // Rejecting the deletion gives the text back whole.
    const list = revisionsOf(o.doc).filter((revision) => revision.kind === 'delete');
    assert.equal(list.length, 1, xml);
    run(o.history, rejectRevision(o.doc, list[0]));
    assert.deepEqual(texts(o.doc), ['xabcdy']);
  }
  assert.equal(new Set(results).size, 1);
});

test('review 4 (A7): where the control keeps something else, all of the own text goes for good, however its runs are cut', async () => {
  for (const inner of [r('abcd'), r('ab') + r('cd', '<w:b/>')]) {
    const o = await open(p(r('x') + CONTROL(r('k') + OWN(inner)) + r('y')), { settings: '<w:trackRevisions/>' });
    run(o.history, deleteRange(o.doc, range(pos(0, 2), pos(0, 6)), ME));
    assert.deepEqual(texts(o.doc), ['xky']);
    const xml = await body(o);
    assert.equal(deletedText(xml), '', xml);
    assert.equal(count(xml, '<w:ins '), 0, xml);
  }
});

test('review 4 (A7): a link inside a control, each of them left empty or not', async () => {
  // The link holds nothing but the own text, the control holds other text too: the text stays deleted, for the link.
  const link = (inner: string) => `<w:hyperlink w:anchor="here">${inner}</w:hyperlink>`;
  for (const inner of [r('abcd'), r('ab') + r('cd', '<w:b/>')]) {
    const o = await open(p(r('x') + CONTROL(r('k') + link(OWN(inner))) + r('y')), { settings: '<w:trackRevisions/>' });
    run(o.history, deleteRange(o.doc, range(pos(0, 2), pos(0, 6)), ME));
    const xml = await body(o);
    assert.equal(deletedText(xml), 'abcd', xml);
    assert.equal(count(xml, '<w:hyperlink '), 1, xml);
  }
});

test('review 4 (A7): Backspace key by key takes own text out of a control for good until the last character, which stays deleted', async () => {
  const o = await open(p(r('x') + CONTROL(OWN(r('abc'))) + r('y')), { settings: '<w:trackRevisions/>' });
  for (let offset = 4; offset > 1; offset--) run(o.history, deleteBackward(o.doc, pos(0, offset), 'char', ME));
  assert.deepEqual(texts(o.doc), ['xy']);
  const xml = await body(o);
  assert.equal(deletedText(xml), 'a', xml);
  assert.equal(count(xml, '<w:id w:val="99"/>'), 1, xml);
});

// ------------------------------------------------------------------------- A8

test('review 4 (A8): rejecting a change of formatting never makes a section break or a tracked change out of what the record holds', async () => {
  const badPara =
    `<w:jc w:val="center"/><w:pPrChange ${WHO(40)}><w:pPr><w:jc w:val="right"/>` +
    `<w:rPr><w:ins ${WHO(41, 'Eve')}/><w:b/></w:rPr><w:sectPr><w:pgSz w:w="3000" w:h="3000"/></w:sectPr><w:pPrChange ${WHO(42, 'Eve')}><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr></w:pPrChange>`;
  const badRun = `<w:b/><w:rPrChange ${WHO(43)}><w:rPr><w:i/><w:ins ${WHO(44, 'Eve')}/><w:del ${WHO(45, 'Eve')}/><w:rPrChange ${WHO(46, 'Eve')}><w:rPr><w:u w:val="single"/></w:rPr></w:rPrChange></w:rPr></w:rPrChange>`;
  const badRow =
    `<w:tr><w:trPr><w:cantSplit/><w:trPrChange ${WHO(47)}><w:trPr><w:ins ${WHO(48, 'Eve')}/><w:del ${WHO(49, 'Eve')}/><w:trPrChange ${WHO(50, 'Eve')}><w:trPr/></w:trPrChange></w:trPr></w:trPrChange></w:trPr>` +
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:tcPrChange ${WHO(51)}><w:tcPr><w:tcW w:w="1000" w:type="dxa"/><w:tcPrChange ${WHO(52, 'Eve')}><w:tcPr/></w:tcPrChange></w:tcPr></w:tcPrChange></w:tcPr>${p('cell')}</w:tc></w:tr>`;
  const o = await open(p(r('para'), badPara) + p(r('run', badRun), `<w:rPr>${badRun}</w:rPr>`) + tbl(badRow, [2000]) + p('end'));
  const sections = (doc: DocModel): number => doc.body.blocks.filter((block) => block.type === 'paragraph' && block.props.section !== undefined).length;
  assert.equal(sections(o.doc), 0);
  // What the file shows of changes: the six records, none of what is inside them.
  const before = revisionsOf(o.doc);
  assert.ok(before.every((revision) => revision.author === 'Ann'), JSON.stringify(before));
  const edit = rejectAll(o.doc);
  assert.equal(edit.skipped, 0);
  run(o.history, edit);
  assert.deepEqual(revisionsOf(o.doc), []);
  assert.equal(sections(o.doc), 0, 'a section break out of a record');
  const first = paragraphAt(o.doc, 0).props;
  assert.equal(first.align, 'right');
  assert.equal(first.markProps, undefined);
  assert.equal(first.change, undefined);
  const second = paragraphAt(o.doc, 1);
  assert.deepEqual(second.props.markProps, { italic: true });
  const inline = second.children[0];
  assert.ok(inline.type === 'run');
  assert.deepEqual(inline.props, { italic: true });
  const table = o.doc.body.blocks[2];
  assert.ok(table.type === 'table' && table.rows.length === 1);
  assert.equal(table.rows[0].props.inserted, undefined);
  assert.equal(table.rows[0].props.deleted, undefined);
  assert.equal(table.rows[0].props.change, undefined);
  assert.equal(table.rows[0].cells[0].props.change, undefined);
  const xml = await written(o);
  parseXml(xml);
  assert.ok(!/Eve/.test(xml.slice(xml.indexOf('<w:body>'))), xml);
  assert.deepEqual(texts(o.doc), ['para', 'run', 'cell', 'end']);
  await assertRereads(o);
  o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);
});

// ------------------------------------------------------------------------- A9

test('review 4 (A9): a block an edit hands back as it was keeps the id its file gave it, also one the file has twice', async () => {
  const gone = `<w:del ${WHO(5, 'Bob')}><w:r><w:delText>gone</w:delText></w:r></w:del>`;
  const o = await open(p(`<w:ins ${WHO(5)}>${r('aaaa')}</w:ins>`) + p(gone, DEL_MARK(6, 'Bob')) + p('tail'), { settings: '<w:trackRevisions/>' });
  const untouched = o.doc.body.blocks[1];
  run(o.history, deleteRange(o.doc, range(pos(0, 1), pos(2, 1)), ME));
  // The paragraph in the middle had nothing left to delete: it is the block it was, and is written from its source.
  assert.equal(o.doc.body.blocks[1], untouched);
  await assertRereads(o);
  const xml = await body(o);
  assert.ok(xml.indexOf(p(gone, DEL_MARK(6, 'Bob'))) > 0, xml);
  // Typing over a selection in such a file, tracked: the same.
  const typed = await open(p(`<w:ins ${WHO(5)}>${r('aaaa')}</w:ins>`) + p(gone, DEL_MARK(6, 'Bob')) + p('tail'), { settings: '<w:trackRevisions/>' });
  run(typed.history, insertText(typed.doc, range(pos(0, 1), pos(2, 1)), 'new', undefined, ME));
  await assertRereads(typed);
  parseXml(await written(typed));
});

test('review 4 (A9): what an edit does change still gets ids no block beside it has', async () => {
  const o = await open(p(`<w:ins ${WHO(5)}>${r('aaaa')}</w:ins>`) + p(`<w:ins ${WHO(5, 'Bob')}>${r('bbbb')}</w:ins>`, INS_MARK(5, 'Cid')) + p('tail'));
  // Not tracked: the second paragraph is cut in two, both halves are new blocks.
  run(o.history, splitParagraph(o.doc, pos(1, 2)));
  await assertRereads(o);
  const xml = await body(o);
  const second = xml.slice(xml.indexOf('</w:p>') + 6);
  assertUniqueIds(second);
});

test('review 4: an item that leaves its list is out of the list in the file too, also when its number was a recorded change', async () => {
  const o = await open(p('', LISTED) + p('tail'), { numbering: NUMBERING });
  // Enter in an empty item of a list takes the item out of the list.
  run(o.history, splitParagraph(o.doc, pos(0, 0)));
  assert.equal(paragraphAt(o.doc, 0).props.numbering, undefined);
  const xml = await body(o);
  assert.equal(count(xml, '<w:numPr>'), 0, xml);
  await assertRereads(o);
  o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);
});

// ------------------------------------------------------------------------ A10

test('review 4 (A10): a control that is one again after its two halves are joined has no stand-in for "empty" beside its content', async () => {
  const o = await open(p(r('x') + CONTROL(`<w:ins ${WHO(7)}>${r('a')}</w:ins>` + r('b')) + r('y')));
  // Enter inside the control, tracked: both paragraphs hold a part of it.
  run(o.history, splitParagraph(o.doc, pos(0, 2), TRACKED));
  assert.deepEqual(texts(o.doc), ['xa', 'by']);
  await assertRereads(o, 'split');
  // Ann's text goes (the first half holds nothing then), and so does the paragraph mark: the halves are one control again.
  const edit = rejectAll(o.doc);
  assert.equal(edit.skipped, 0);
  run(o.history, edit);
  assert.deepEqual(texts(o.doc), ['xby']);
  assert.deepEqual(
    paragraphAt(o.doc, 0).children.map((inline) => inline.type),
    ['run', 'run', 'run'],
  );
  await assertRereads(o, 'rejected');
  const xml = await body(o);
  assert.equal(count(xml, '<w:sdt>'), 1, xml);
  // One by one, the mark first, then the text: the control is left with "b", and no stand-in.
  o.history.undo();
  const list = revisionsOf(o.doc);
  assert.deepEqual(
    list.map((revision) => revision.kind),
    ['insert', 'insert'],
  );
  run(o.history, rejectRevision(o.doc, list[1]));
  await assertRereads(o, 'mark rejected');
  run(o.history, rejectRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.deepEqual(texts(o.doc), ['xby']);
  assert.ok(!paragraphAt(o.doc, 0).children.some((inline) => inline.type === 'empty'));
  await assertRereads(o, 'both rejected');
  // The other order: the text first leaves the first half empty, with a stand-in; the join takes it away.
  o.history.undo();
  o.history.undo();
  run(o.history, rejectRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.ok(paragraphAt(o.doc, 0).children.some((inline) => inline.type === 'empty'));
  await assertRereads(o, 'text rejected');
  run(o.history, rejectRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.ok(!paragraphAt(o.doc, 0).children.some((inline) => inline.type === 'empty'));
  await assertRereads(o, 'text, then mark');
  while (o.history.canUndo) o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);
});

test('review 4 (A10): an empty control of the file keeps its stand-in, and loses it when a join without tracking fills it', async () => {
  const o = await open(p(r('x') + CONTROL(r('ab')) + r('y')));
  run(o.history, splitParagraph(o.doc, pos(0, 2)));
  // The first half is emptied: its stand-in is made by nobody, the file is written without content there.
  run(o.history, deleteBackward(o.doc, pos(0, 2)));
  await assertRereads(o, 'emptied');
  run(o.history, deleteForward(o.doc, pos(0, 1)));
  assert.deepEqual(texts(o.doc), ['xby']);
  await assertRereads(o, 'joined');
});
