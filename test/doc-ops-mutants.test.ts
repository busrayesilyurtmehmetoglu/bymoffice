// Tests for branches of the document editing code that mutation testing found
// no test for: each of them fails when the branch it names is changed.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { paragraphText, type Block, type DocModel, type Paragraph, type Run, type Table, type Wrapper } from '../src/model/doc/types';
import {
  DocHistory,
  applyDocOp,
  comparePositions,
  deleteColumn,
  deleteForward,
  deleteRow,
  insertColumn,
  insertHyperlink,
  insertImage,
  insertPageBreak,
  insertRow,
  insertText,
  movePosition,
  patchRunProps,
  setRunFormat,
  splitParagraph,
  storyOf,
  type DocOp,
  type DocRange,
} from '../src/ops/doc/index';
import { flatten, removePieces, wrapBetween } from '../src/ops/doc/pieces';
import { docPicture, open, p, paragraphAt, pos, r, range, run, shapeAt, tbl, tc, texts, tr } from './doc-ops-fixtures';

const history = (doc: DocModel) => new DocHistory<DocRange | undefined>(doc);
const para = (text: string): Paragraph => ({ type: 'paragraph', props: {}, children: text ? [{ type: 'run', props: {}, content: [{ type: 'text', text }] }] : [] });
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3]);

test('a new relationship takes an id that is free, counting ids as numbers', async () => {
  const link = (doc: DocModel): string => {
    const before = new Set(doc.body.relationships.keys());
    run(history(doc), insertHyperlink(doc, range(pos(0, 0), pos(0, 2)), 'https://example.org/'));
    const added = Array.from(doc.body.relationships.keys()).filter((id) => !before.has(id));
    assert.equal(added.length, 1);
    return added[0];
  };
  const one = await open({ body: p('text'), rels: [['rId1', 'hyperlink', 'https://a.example/', true]] });
  assert.equal(link(one), 'rId2');
  // Ten comes after nine, and an id that is not of the form does not count.
  const ten = await open({ body: p('text'), rels: [['rId9', 'hyperlink', 'https://a.example/', true], ['rId10', 'hyperlink', 'https://b.example/', true], ['rIdX77', 'hyperlink', 'https://c.example/', true], ['other99', 'hyperlink', 'https://d.example/', true]] });
  assert.equal(link(ten), 'rId11');
  // A picture and a link made one after the other do not share an id.
  const both = await open({ body: p('text'), rels: [['rId1', 'hyperlink', 'https://a.example/', true]] });
  const h = history(both);
  run(h, insertImage(both, pos(0, 0), { bytes: PNG, type: 'image/png', width: 9525, height: 9525 }));
  run(h, insertHyperlink(both, range(pos(0, 0), pos(0, 2)), 'https://example.org/'));
  assert.equal(both.body.relationships.get('rId1')!.target, 'https://a.example/');
  assert.deepEqual(Array.from(both.body.relationships.keys()).filter((id) => /^rId\d+$/.test(id)).sort(), ['rId1', 'rId2', 'rId3']);
});

test('the index of a section mark in the middle of the body follows every kind of change, and comes back', async () => {
  const SECTION = '<w:sectPr><w:pgSz w:w="9000" w:h="9000"/></w:sectPr>';
  const doc = await open({ body: p('one') + p('two') + SECTION + p('three') + p('four') });
  assert.equal(doc.body.sectionIndex, 2);
  const apply = (op: DocOp): DocOp => applyDocOp(doc, op);
  // A replacement that ends before the mark shifts it by the change in the number of blocks.
  let inverse = apply({ type: 'replaceBlocks', story: 'body', path: [], at: 0, remove: 2, insert: [para('a')] });
  assert.equal(doc.body.sectionIndex, 1);
  assert.equal((inverse as { sectionIndex?: number }).sectionIndex, 2);
  apply(inverse);
  assert.equal(doc.body.sectionIndex, 2);
  // One that reaches over the mark puts it after what was inserted.
  inverse = apply({ type: 'replaceBlocks', story: 'body', path: [], at: 1, remove: 2, insert: [para('x'), para('y'), para('z')] });
  assert.equal(doc.body.sectionIndex, 4);
  apply(inverse);
  assert.equal(doc.body.sectionIndex, 2);
  assert.deepEqual(texts(doc), ['one', 'two', 'three', 'four']);
  // One that begins at the mark, or after it, leaves it alone, and its inverse says nothing about it.
  inverse = apply({ type: 'replaceBlocks', story: 'body', path: [], at: 2, remove: 1, insert: [] });
  assert.equal(doc.body.sectionIndex, 2);
  assert.equal('sectionIndex' in inverse, false);
  apply(inverse);
  // The index can be given, or taken away, and both come back.
  inverse = apply({ type: 'replaceBlocks', story: 'body', path: [], at: 0, remove: 0, insert: [], sectionIndex: 4 });
  assert.equal(doc.body.sectionIndex, 4);
  apply(inverse);
  assert.equal(doc.body.sectionIndex, 2);
  inverse = apply({ type: 'replaceBlocks', story: 'body', path: [], at: 0, remove: 0, insert: [], sectionIndex: null });
  assert.equal('sectionIndex' in doc.body, false);
  const again = apply(inverse);
  assert.equal(doc.body.sectionIndex, 2);
  assert.equal((again as { sectionIndex?: number | null }).sectionIndex, null);
  // A body without the index gets none from a change, and its inverse names none.
  const plain = await open(p('one') + p('two'));
  assert.equal(plain.body.sectionIndex, undefined);
  const undo = applyDocOp(plain, { type: 'replaceBlocks', story: 'body', path: [], at: 0, remove: 1, insert: [para('a'), para('b')] });
  assert.equal('sectionIndex' in plain.body, false);
  assert.equal('sectionIndex' in undo, false);
  // Other stories have no section mark.
  const noted = await open({ body: p('x'), headers: { rIdH: p('head') } });
  assert.equal('sectionIndex' in applyDocOp(noted, { type: 'replaceBlocks', story: 'body', path: [], at: 0, remove: 0, insert: [para('n')], sectionIndex: 7 }), true);
});

function manyParagraphs(count: number): Block[] {
  const out: Block[] = [];
  const shared = para('x');
  for (let i = 0; i < count; i++) out.push({ ...shared });
  return out;
}

test('undo steps go by what they hold: the oldest are dropped beyond the limit, the newest always stays', async () => {
  const doc = await open(p('x'));
  const SIZE = 50000;
  doc.body.blocks = manyParagraphs(SIZE);
  const h = new DocHistory(doc);
  const all = (): DocOp => ({ type: 'replaceBlocks', story: 'body', path: [], at: 0, remove: SIZE, insert: manyParagraphs(SIZE) });
  // Each step holds 50,000 blocks; the limit is 400,000.
  for (let i = 0; i < 10; i++) h.apply([all()], undefined, undefined);
  let undone = 0;
  while (h.undo()) undone++;
  assert.equal(undone, 7);
  // Undoing gave the weight back: seven more steps fit again, and redo counts too.
  for (let i = 0; i < 7; i++) h.apply([all()], undefined, undefined);
  for (let i = 0; i < 3; i++) h.undo();
  for (let i = 0; i < 3; i++) h.redo();
  h.apply([all()], undefined, undefined);
  undone = 0;
  while (h.undo()) undone++;
  assert.equal(undone, 7);
  // A single step beyond the limit is kept: there is always something to undo.
  const big = await open(p('x'));
  big.body.blocks = manyParagraphs(450000);
  const first = big.body.blocks[0];
  const one = new DocHistory(big);
  one.apply([{ type: 'replaceBlocks', story: 'body', path: [], at: 0, remove: 450000, insert: [para('all gone')] }], undefined, undefined);
  assert.equal(one.canUndo, true);
  one.undo();
  assert.equal(big.body.blocks.length, 450000);
  assert.equal(big.body.blocks[0], first);
});

function bigTable(rows: number, columns: number): Table {
  const list = [];
  for (let i = 0; i < rows; i++) {
    const cells = [];
    for (let c = 0; c < columns; c++) cells.push({ props: {}, blocks: [para('cell')] as Block[] });
    list.push({ props: {}, cells });
  }
  const grid: number[] = [];
  for (let c = 0; c < columns; c++) grid.push(1000);
  return { type: 'table', props: {}, grid, rows: list };
}

test('a step weighs what it keeps alive: a cell of a large table is light, a table that was removed is heavy', async () => {
  const doc = await open(p('x'));
  doc.body.blocks.splice(0, 0, bigTable(20000, 2));
  const h = history(doc);
  // 150 changes of one cell each hold a row and a cell, not the table: none of them is dropped.
  for (let i = 0; i < 150; i++) run(h, insertText(doc, pos([0, i * 7, 1, 0], 2), 'word'));
  let undone = 0;
  while (h.undo()) undone++;
  assert.equal(undone, 150);
  // Eight steps that each remove a table of 60,000 cells hold 480,000 paragraphs: the first ones are dropped.
  const heavy = await open(p('x'));
  const tables: Block[] = [];
  for (let i = 0; i < 8; i++) tables.push(bigTable(600, 100));
  heavy.body.blocks.splice(0, 0, ...tables);
  const h2 = new DocHistory(heavy);
  for (let i = 0; i < 8; i++) h2.apply([{ type: 'replaceBlocks', story: 'body', path: [], at: 0, remove: 1, insert: [] }], undefined, undefined);
  undone = 0;
  while (h2.undo()) undone++;
  assert.ok(undone >= 5 && undone < 8, `${undone} of 8 steps could be undone`);
});

test('typing merged into a step: the same block keeps the first inverse, another block or shape adds to the step', async () => {
  const doc = await open(p('one') + p('two') + p('three'));
  const [one, two, three] = doc.body.blocks;
  const h = new DocHistory(doc);
  const replace = (at: number, text: string): DocOp => ({ type: 'replaceBlocks', story: 'body', path: [], at, remove: 1, insert: [para(text)] });
  h.apply([replace(0, 'ONE')], undefined, undefined, 'key');
  h.apply([replace(0, 'ONE!')], undefined, undefined, 'key');
  // Same key, another block: both changes are one step, and undoing takes both back.
  h.apply([replace(1, 'TWO')], undefined, undefined, 'key');
  // Same key, two operations at once.
  h.apply([replace(2, 'THREE'), replace(1, 'TWO!')], undefined, undefined, 'key');
  assert.deepEqual(texts(doc), ['ONE!', 'TWO!', 'THREE']);
  assert.deepEqual(h.undo(), { meta: undefined });
  assert.deepEqual(doc.body.blocks, [one, two, three]);
  assert.equal(doc.body.blocks[0], one);
  assert.equal(h.canUndo, false);
  assert.equal('modified' in doc.body, false);
  h.redo();
  // Redo gives the state before the undo.
  assert.deepEqual(texts(doc), ['ONE!', 'TWO!', 'THREE']);
  // Another key, or no key, is another step.
  h.apply([replace(0, 'a')], undefined, undefined, 'other');
  h.apply([replace(0, 'b')], undefined, undefined);
  h.apply([replace(0, 'c')], undefined, undefined, 'other');
  h.undo();
  assert.equal(texts(doc)[0], 'b');
  h.undo();
  assert.equal(texts(doc)[0], 'a');
  h.undo();
  assert.equal(texts(doc)[0], 'ONE!');
  // The selection of a merged step: before from its first part, after from its last.
  const merged = new DocHistory<string>(doc);
  merged.apply([replace(0, 'x')], 'before 1', 'after 1', 'k');
  merged.apply([replace(0, 'xy')], 'before 2', 'after 2', 'k');
  assert.deepEqual(merged.undo(), { meta: 'before 1' });
  assert.deepEqual(merged.redo(), { meta: 'after 2' });
});

test('Enter, a paste of lines and a page break: the new paragraphs do not take the identity of the old one; runs keep theirs', async () => {
  const body = '<w:p w14:paraId="1234ABCD" w:rsidR="00AA0001"><w:r w:rsidRPr="00BB0002"><w:t>headtail</w:t></w:r></w:p>';
  const doc = await open(body);
  const original = paragraphAt(doc, 0);
  assert.ok(original.attrs && original.attrs.length === 2);
  const runAttrs = (original.children[0] as Run).attrs;
  assert.ok(runAttrs && runAttrs.length === 1);
  const h = history(doc);
  run(h, splitParagraph(doc, pos(0, 4)));
  assert.deepEqual(paragraphAt(doc, 0).attrs, original.attrs);
  assert.equal(paragraphAt(doc, 1).attrs, undefined);
  // Both halves of the run are still that run.
  assert.deepEqual((paragraphAt(doc, 0).children[0] as Run).attrs, runAttrs);
  assert.deepEqual((paragraphAt(doc, 1).children[0] as Run).attrs, runAttrs);
  h.undo();
  run(h, insertText(doc, pos(0, 4), 'a\nb\nc'));
  assert.deepEqual([0, 1, 2].map((i) => paragraphAt(doc, i).attrs !== undefined), [true, false, false]);
  assert.deepEqual((paragraphAt(doc, 0).children[0] as Run).attrs, runAttrs, 'text typed into a run stays in it');
  h.undo();
  run(h, insertPageBreak(doc, pos(0, 4)));
  assert.deepEqual([0, 1].map((i) => paragraphAt(doc, i).attrs !== undefined), [true, false]);
  h.undo();
  // Typing does not take the attributes of the paragraph away.
  run(h, insertText(doc, pos(0, 4), '!'));
  assert.deepEqual(paragraphAt(doc, 0).attrs, original.attrs);
});

const cellTable = (text: string) => tbl(tr(tc(text)), [1000]);

test('a table that is removed leaves a paragraph exactly where the list of blocks needs one', async () => {
  const types = (doc: DocModel) => doc.body.blocks.map((block) => block.type);
  const remove = async (body: string, at: number): Promise<DocModel> => {
    const doc = await open(body);
    run(history(doc), deleteRow(doc, pos([at, 0, 0, 0], 0)));
    return doc;
  };
  // Between two tables: they must not touch.
  assert.deepEqual(types(await remove(cellTable('a') + cellTable('b') + cellTable('c') + p('end'), 1)), ['table', 'paragraph', 'table', 'paragraph']);
  // Before a table, with nothing or a paragraph before it: nothing is needed.
  assert.deepEqual(types(await remove(cellTable('a') + cellTable('b') + p('end'), 0)), ['table', 'paragraph']);
  assert.deepEqual(types(await remove(p('x') + cellTable('a') + cellTable('b') + p('end'), 1)), ['paragraph', 'table', 'paragraph']);
  // Between paragraphs, and after a paragraph at the end.
  assert.deepEqual(types(await remove(p('x') + cellTable('a') + p('end'), 1)), ['paragraph', 'paragraph']);
  // At the end of its list after a table: the list must not end with a table.
  const nested = await open(tbl(tr(tc(cellTable('a') + cellTable('b'))), [1000]) + p('end'));
  const caret = run(history(nested), deleteRow(nested, pos([0, 0, 0, 1, 0, 0, 0], 0)));
  assert.deepEqual((nested.body.blocks[0] as Table).rows[0].cells[0].blocks.map((block) => block.type), ['table', 'paragraph']);
  assert.deepEqual(caret.focus, pos([0, 0, 0, 1], 0));
  // Alone in its list.
  const alone = await open(tbl(tr(tc(cellTable('a'))), [1000]) + p('end'));
  run(history(alone), deleteRow(alone, pos([0, 0, 0, 0, 0, 0, 0], 0)));
  assert.deepEqual((alone.body.blocks[0] as Table).rows[0].cells[0].blocks.map((block) => block.type), ['paragraph']);
  // The paragraph that takes the place of a table in a content control is in the control.
  const wrapped = await open(`<w:sdt><w:sdtPr><w:alias w:val="c"/></w:sdtPr><w:sdtContent>${cellTable('a')}</w:sdtContent></w:sdt>` + cellTable('b') + p('end'));
  assert.equal(wrapped.body.blocks[0].wrap!.kind, 'sdt');
  const after = run(history(wrapped), deleteRow(wrapped, pos([0, 0, 0, 0], 0)));
  assert.deepEqual(types(wrapped), ['table', 'paragraph']);
  assert.deepEqual(after.focus, pos([0, 0, 0, 0], 0), 'the caret goes to the paragraph after, in the next table');
  // The caret goes to the end of the paragraph before when nothing follows.
  const last = await open(p('before') + cellTable('a'));
  const end = run(history(last), deleteRow(last, pos([1, 0, 0, 0], 0)));
  assert.deepEqual(end.focus, pos(0, 6));
  const filled = await open(cellTable('a') + `<w:sdt><w:sdtPr><w:alias w:val="c"/></w:sdtPr><w:sdtContent>${cellTable('b')}</w:sdtContent></w:sdt>`);
  run(history(filled), deleteRow(filled, pos([1, 0, 0, 0], 0)));
  assert.deepEqual(types(filled), ['table', 'paragraph']);
  assert.equal(filled.body.blocks[1].wrap!.kind, 'sdt');
});

test('a new row does not repeat as a header below one, and is not a tracked change', async () => {
  const header = '<w:tr><w:trPr><w:tblHeader/><w:cantSplit/><w:ins w:id="5" w:author="A" w:date="2020-01-01T00:00:00Z"/></w:trPr>' + tc('h1', '<w:tcW w:w="1000" w:type="dxa"/><w:tcPrChange w:id="6" w:author="A"><w:tcPr><w:tcW w:w="900" w:type="dxa"/></w:tcPr></w:tcPrChange>') + '</w:tr>';
  const doc = await open(tbl(header + tr(tc('a')), [1000]) + p('end'));
  const table = () => doc.body.blocks[0] as Table;
  const source = table().rows[0];
  assert.equal(source.props.header, true);
  assert.ok(source.props.inserted, 'the row is a tracked insertion');
  assert.ok(source.cells[0].props.change, 'the cell has a tracked change of its properties');
  const h = history(doc);
  run(h, insertRow(doc, pos([0, 0, 0, 0], 0), 'below'));
  let made = table().rows[1];
  assert.equal(made.props.header, undefined, 'only the rows at the top repeat');
  assert.equal(made.props.cantSplit, true, 'the rest of the formatting is taken over');
  assert.equal(made.props.inserted, undefined);
  assert.equal(made.cells[0].props.change, undefined);
  assert.deepEqual(made.cells[0].props.width, { type: 'dxa', value: 1000 });
  h.undo();
  run(h, insertRow(doc, pos([0, 0, 0, 0], 0), 'above'));
  made = table().rows[0];
  assert.equal(made.props.header, true, 'above a header row the new row is a header row');
  assert.equal(made.props.inserted, undefined);
  assert.equal(made.cells[0].props.change, undefined);
  // The row it was made from is not touched.
  assert.equal(table().rows[1], source);
});

test('columns in rows that begin or end beside the grid, and beside horizontally merged cells', async () => {
  const row2 = '<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>' + tc('b2') + tc('c2') + '</w:tr>';
  const row3 = '<w:tr><w:trPr><w:gridAfter w:val="1"/></w:trPr>' + tc('a3') + tc('b3') + '</w:tr>';
  const doc = await open(tbl(tr(tc('a1') + tc('b1', '<w:hMerge w:val="restart"/><w:gridSpan w:val="1"/>') + tc('c1')) + row2 + row3, [1000, 1000, 1000]) + p('end'));
  const table = () => doc.body.blocks[0] as Table;
  const edges = () => table().rows.map((row) => [row.props.gridBefore || 0, row.cells.length, row.props.gridAfter || 0]);
  assert.deepEqual(edges(), [[0, 3, 0], [1, 2, 0], [0, 2, 1]]);
  assert.equal(table().rows[0].cells[1].props.hMerge, 'restart');
  const h = history(doc);
  // Left of the first column: the row that begins one column in begins two columns in.
  run(h, insertColumn(doc, pos([0, 0, 0, 0], 0), 'left'));
  assert.deepEqual(edges(), [[0, 4, 0], [2, 2, 0], [0, 3, 1]]);
  h.undo();
  // Right of the last column: the row that ends one column early ends two columns early.
  run(h, insertColumn(doc, pos([0, 0, 2, 0], 0), 'right'));
  assert.deepEqual(edges(), [[0, 4, 0], [1, 3, 0], [0, 2, 2]]);
  h.undo();
  // Beside a horizontally merged cell: the new cell is not part of the merge.
  run(h, insertColumn(doc, pos([0, 0, 1, 0], 0), 'right'));
  assert.deepEqual(edges(), [[0, 4, 0], [1, 3, 0], [0, 3, 1]]);
  assert.equal(table().rows[0].cells[2].props.hMerge, undefined);
  assert.equal(table().rows[0].cells[2].props.gridSpan, undefined);
  assert.equal(table().rows[0].cells[1].props.hMerge, 'restart');
  h.undo();
  // Deleting the first column takes the column the second row leaves out; deleting the last, the one the third leaves out.
  run(h, deleteColumn(doc, pos([0, 0, 0, 0], 0)));
  assert.deepEqual(edges(), [[0, 2, 0], [0, 2, 0], [0, 1, 1]]);
  assert.equal('gridBefore' in table().rows[1].props ? table().rows[1].props.gridBefore : 0, 0);
  h.undo();
  run(h, deleteColumn(doc, pos([0, 0, 2, 0], 0)));
  assert.deepEqual(edges(), [[0, 2, 0], [1, 1, 0], [0, 2, 0]]);
  h.undo();
  assert.deepEqual(edges(), [[0, 3, 0], [1, 2, 0], [0, 2, 1]]);
});

const fieldMarks = (code: string, result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

test('removePieces: a field without a result, and one whose result cannot be removed, keep their marks', async () => {
  const kinds = (paragraph: Paragraph, from: number, to: number): string[] =>
    removePieces(flatten(paragraph), from, to).map((piece) => (piece.item ? (piece.item.type === 'text' ? piece.item.text : piece.item.type === 'fieldChar' ? piece.item.charType : piece.item.type) : piece.inline!.type));
  // "a", begin, code, separate, end, "b": the range takes the separator, the end and "b", but not the beginning.
  const empty = await open(p(r('a') + fieldMarks('REF x', '') + r('b')));
  assert.deepEqual(kinds(paragraphAt(empty, 0), 3, 6), ['a', 'begin', 'instrText', 'separate', 'end']);
  // All of it in the range: the field goes.
  assert.deepEqual(kinds(paragraphAt(empty, 0), 1, 5), ['a', 'b']);
  // A result with text: when all of the result is in the range the field goes, whatever of its marks is.
  const filled = await open(p(r('a') + fieldMarks('REF x', r('res')) + r('b')));
  assert.deepEqual(kinds(paragraphAt(filled, 0), 4, 5), ['a', 'b']);
  assert.deepEqual(kinds(paragraphAt(filled, 0), 0, 1), ['begin', 'instrText', 'separate', 'res', 'end', 'b']);
  // A result that is written back whole cannot be removed: the field stays around it.
  const alternate = '<mc:AlternateContent><mc:Choice Requires="wps"><w:r><w:t>choice</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>fb</w:t></w:r></mc:Fallback></mc:AlternateContent>';
  const frozen = await open(p(r('a') + fieldMarks('REF x', alternate) + r('b')));
  const pieces = flatten(paragraphAt(frozen, 0));
  const held = pieces.findIndex((piece) => piece.frozen === true);
  assert.ok(held > 0 && pieces[held].len > 0);
  const left = kinds(paragraphAt(frozen, 0), held, pieces.length);
  assert.deepEqual(left.slice(0, 4), ['a', 'begin', 'instrText', 'separate']);
  assert.equal(left[left.length - 1], 'end');
  assert.equal(left.indexOf('b'), -1);
});

test('new content goes outside of what is not for new text: deletions, alternate content, computed fields', () => {
  const link: Wrapper = { kind: 'hyperlink', anchor: 'x' };
  const page: Wrapper = { kind: 'simpleField', instruction: ' PAGE ', parent: link };
  const author: Wrapper = { kind: 'simpleField', instruction: ' AUTHOR ', parent: link };
  const deleted: Wrapper = { kind: 'del', parent: link };
  const moved: Wrapper = { kind: 'moveFrom', parent: link };
  const inserted: Wrapper = { kind: 'ins', parent: link };
  const alternate: Wrapper = { kind: 'alternate', selected: 'choice', xml: { type: 'element', name: 'mc:AlternateContent', attrs: new Map(), children: [] }, parent: link };
  assert.equal(wrapBetween(page, page), link, 'not into a field the engine computes');
  assert.equal(wrapBetween(author, author), author, 'into the result of another field');
  assert.equal(wrapBetween(deleted, deleted), link);
  assert.equal(wrapBetween(moved, moved), link);
  assert.equal(wrapBetween(alternate, alternate), link);
  assert.equal(wrapBetween(inserted, inserted), inserted);
  assert.equal(wrapBetween(page, author), link, 'two neighbours share what is around both');
  assert.equal(wrapBetween(page, undefined), undefined);
  assert.equal(wrapBetween(undefined, link), undefined);
  const inner: Wrapper = { kind: 'ins', parent: page };
  assert.equal(wrapBetween(inner, inner), link, 'an insertion inside a computed field is left with the field');
});

test('content that is written back whole refuses edits wherever it sits: around a table, a row, a cell, a block of a cell', async () => {
  const xml = { type: 'element' as const, name: 'mc:AlternateContent', attrs: new Map<string, string>(), children: [] };
  const whole = (): Wrapper => ({ kind: 'alternate', selected: 'choice', xml });
  const make = async (): Promise<DocModel> => open(p('before') + tbl(tr(tc('a1') + tc(p('b1') + cellTable('inner') + p('b1 end'))) + tr(tc('a2') + tc('b2')), [1000, 1000]) + p('after'));
  const refuses = (doc: DocModel, at: number[], message: string) => {
    const picture = docPicture(doc);
    assert.throws(() => insertText(doc, pos(at, 0), 'x'), RangeError, message);
    assert.throws(() => splitParagraph(doc, pos(at, 0)), RangeError, message);
    assert.throws(() => setRunFormat(doc, range(pos(at, 0), pos(at, 1)), { bold: true }), RangeError, message);
    assert.equal(docPicture(doc), picture, message);
  };
  let doc = await make();
  (doc.body.blocks[1] as Table).wrap = whole();
  refuses(doc, [1, 0, 0, 0], 'the table');
  refuses(doc, [1, 0, 1, 1, 0, 0, 0], 'a table in the table');
  doc = await make();
  (doc.body.blocks[1] as Table).rows[0].wrap = whole();
  refuses(doc, [1, 0, 0, 0], 'the row');
  assert.equal(insertText(doc, pos([1, 1, 0, 0], 0), 'x').ops.length, 1, 'the other row can be edited');
  doc = await make();
  (doc.body.blocks[1] as Table).rows[0].cells[1].wrap = whole();
  refuses(doc, [1, 0, 1, 0], 'the cell');
  refuses(doc, [1, 0, 1, 1, 0, 0, 0], 'a table in the cell');
  assert.equal(insertText(doc, pos([1, 0, 0, 0], 0), 'x').ops.length, 1, 'the cell beside it can be edited');
  doc = await make();
  (doc.body.blocks[1] as Table).rows[0].cells[1].blocks[0].wrap = whole();
  refuses(doc, [1, 0, 1, 0], 'a paragraph of the cell');
  assert.equal(insertText(doc, pos([1, 0, 1, 2], 0), 'x').ops.length, 1, 'the paragraph after it can be edited');
  doc = await make();
  ((doc.body.blocks[1] as Table).rows[0].cells[1].blocks[1] as Table).rows[0].cells[0].wrap = whole();
  refuses(doc, [1, 0, 1, 1, 0, 0, 0], 'a cell of the inner table');
  // Removing the row or the table that holds such content is refused too: it could not be put back together.
  doc = await make();
  (doc.body.blocks[1] as Table).rows[0].cells[1].blocks[0].wrap = whole();
  const picture = docPicture(doc);
  assert.throws(() => deleteRow(doc, pos([1, 0, 0, 0], 0)), RangeError);
  assert.equal(docPicture(doc), picture);
});

const BOX = (inner: string) =>
  '<w:r><w:drawing><wp:inline><wp:extent cx="190500" cy="190500"/><wp:docPr id="7" name="s"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx><w:txbxContent>' +
  inner +
  '</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';

test('positions in a text box compare with positions of the same box only', async () => {
  const doc = await open(p(r('text') + BOX(p('one') + p('two')) + BOX(p('other'))) + p('more'));
  const first = [0, -1, 0, 0];
  const second = [0, -1, 1, 0];
  assert.equal(paragraphText(paragraphAt(doc, first.concat([1]))), 'two');
  assert.equal(paragraphText(paragraphAt(doc, second.concat([0]))), 'other');
  assert.ok(comparePositions(pos(first.concat([0]), 2), pos(first.concat([1]), 0)) < 0);
  assert.ok(comparePositions(pos(first.concat([1]), 0), pos(first.concat([0]), 2)) > 0);
  assert.ok(comparePositions(pos(first.concat([1]), 1), pos(first.concat([1]), 3)) < 0);
  assert.equal(comparePositions(pos(first.concat([1]), 1), pos(first.concat([1]), 1)), 0);
  assert.throws(() => comparePositions(pos(first.concat([0]), 0), pos(second.concat([0]), 0)), RangeError, 'another box');
  assert.throws(() => comparePositions(pos(0, 0), pos(first.concat([0]), 0)), RangeError, 'the paragraph the box is in');
  assert.throws(() => comparePositions(pos(first.concat([0]), 0), pos(1, 0)), RangeError, 'a paragraph after the box');
  assert.throws(() => comparePositions(pos([0, -1, 0, 1, 0], 0), pos(first.concat([0]), 0)), RangeError, 'another box of the same drawing');
  // Moving stays inside the box.
  assert.deepEqual(movePosition(doc, pos(first.concat([1]), 3), 'char', 1), pos(first.concat([1]), 3));
  assert.deepEqual(movePosition(doc, pos(first.concat([0]), 0), 'char', -1), pos(first.concat([0]), 0));
  assert.deepEqual(movePosition(doc, pos(first.concat([0]), 3), 'char', 1), pos(first.concat([1]), 0));
  assert.deepEqual(movePosition(doc, pos(first.concat([0]), 1), 'story', 1), pos(first.concat([1]), 3));
  assert.deepEqual(movePosition(doc, pos(first.concat([1]), 1), 'story', -1), pos(first.concat([0]), 0));
});

test('moving back by paragraphs: to the start of this one first, then to the start of the one before', async () => {
  const doc = await open(p('first') + tbl(tr(tc(p('c1') + p('c2'))), [1000]) + p('last'));
  assert.deepEqual(movePosition(doc, pos(2, 3), 'paragraph', -1), pos(2, 0));
  assert.deepEqual(movePosition(doc, pos(2, 0), 'paragraph', -1), pos([1, 0, 0, 1], 0));
  assert.deepEqual(movePosition(doc, pos([1, 0, 0, 1], 0), 'paragraph', -1), pos([1, 0, 0, 0], 0));
  assert.deepEqual(movePosition(doc, pos([1, 0, 0, 0], 1), 'paragraph', -1), pos([1, 0, 0, 0], 0));
  assert.deepEqual(movePosition(doc, pos([1, 0, 0, 0], 0), 'paragraph', -1), pos(0, 0));
  assert.deepEqual(movePosition(doc, pos(0, 0), 'paragraph', -1), pos(0, 0), 'at the start of the story nothing moves');
  assert.deepEqual(movePosition(doc, pos(2, 1), 'paragraph', 1), pos(2, 4), 'in the last paragraph: to its end');
  assert.deepEqual(movePosition(doc, pos(0, 2), 'paragraph', 1), pos([1, 0, 0, 0], 0));
});

test('the id of a note in the name of its story is a plain whole number', async () => {
  const doc = await open({
    body: p(r('x') + '<w:r><w:footnoteReference w:id="2"/></w:r>'),
    footnotes: '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:footnote>',
  });
  assert.equal(storyOf(doc, 'footnote:2').noteId, 2);
  assert.equal(storyOf(doc, 'footnote:-1').noteId, -1);
  for (const id of ['footnote:2x', 'footnote:x2', 'footnote: 2', 'footnote:2 ', 'footnote:2.0', 'footnote:1e0', 'footnote:0x2', 'footnote:+2', 'footnote:--1', 'footnote:', 'footnote', 'footnote:0000000002', 'footnote:2:2', 'endnote:2', 'footnotes:2', 'Footnote:2']) {
    assert.throws(() => storyOf(doc, id), RangeError, id);
  }
  assert.throws(() => storyOf(doc, 2 as unknown as string), RangeError);
});

test('numbering: an id given twice in one operation ends as the last says, and undoing gives back what was there', async () => {
  const doc = await open(p('x'));
  const numbering = doc.numbering;
  const first = numbering.nums.get(1)!;
  const made = (abstractId: number) => ({ id: 1, abstractId, overrides: new Map() });
  const a = made(0);
  const b = made(1);
  let inverse = applyDocOp(doc, { type: 'setNumbering', nums: [[1, a], [1, b]] });
  assert.equal(numbering.nums.get(1), b);
  applyDocOp(doc, inverse);
  assert.equal(numbering.nums.get(1), first, 'the instance that was there, not the one of the first entry');
  assert.equal('modified' in numbering, false);
  // A new id set and removed in one operation; an id set twice that was not there.
  inverse = applyDocOp(doc, { type: 'setNumbering', nums: [[9, { ...a, id: 9 }], [9, null]], abstracts: [[7, { id: 7, levels: new Map() }], [7, { id: 7, levels: new Map(), multiLevelType: 'x' }]] });
  assert.equal(numbering.nums.has(9), false);
  assert.equal(numbering.abstracts.get(7)!.multiLevelType, 'x');
  assert.equal(numbering.modified, true);
  applyDocOp(doc, inverse);
  assert.equal(numbering.nums.has(9), false);
  assert.equal(numbering.abstracts.has(7), false);
  assert.equal(numbering.nums.size, 2);
  // Removing and setting again: undone, the first one is back.
  inverse = applyDocOp(doc, { type: 'setNumbering', nums: [[1, null], [1, a]] });
  assert.equal(numbering.nums.get(1), a);
  applyDocOp(doc, inverse);
  assert.equal(numbering.nums.get(1), first);
});

test('striking through: a single line replaces a double one; switching it off or removing it does not', () => {
  const doubled = { doubleStrike: true, bold: true };
  assert.deepEqual(patchRunProps(doubled, { strike: true }), { bold: true, strike: true });
  assert.deepEqual(patchRunProps(doubled, { strike: false }), { doubleStrike: true, bold: true, strike: false });
  assert.equal(patchRunProps(doubled, { strike: null }), doubled, 'nothing to remove: the same object');
  assert.deepEqual(patchRunProps({ strike: true, doubleStrike: true }, { strike: null }), { doubleStrike: true });
  assert.equal(patchRunProps(doubled, { italic: null }), doubled);
});

test('the formatting of a paragraph mark: gone when its last property goes; typed text does not take its revision marks', async () => {
  const doc = await open(p('one', '<w:rPr><w:b/><w:bCs/></w:rPr>') + p('two', '<w:jc w:val="center"/><w:rPr><w:b/><w:bCs/></w:rPr>') + p('three') + p('', '<w:rPr><w:ins w:id="1" w:author="A"/><w:b/></w:rPr>') + p('', '<w:rPr><w:ins w:id="2" w:author="A"/></w:rPr>'));
  const h = history(doc);
  assert.deepEqual(paragraphAt(doc, 0).props, { markProps: { bold: true, boldCs: true } });
  run(h, setRunFormat(doc, range(pos(0, 0), pos(2, 2)), { bold: null }));
  // No empty bag is left behind, and a paragraph without properties shares the empty one.
  assert.equal('markProps' in paragraphAt(doc, 0).props, false);
  assert.deepEqual(paragraphAt(doc, 0).props, {});
  assert.deepEqual(paragraphAt(doc, 1).props, { align: 'center' });
  // The mark of the last paragraph of the range is not in the range.
  run(h, setRunFormat(doc, range(pos(0, 0), pos(1, 3)), { italic: true }));
  assert.deepEqual(paragraphAt(doc, 0).props.markProps, { italic: true, italicCs: true });
  assert.equal(paragraphAt(doc, 1).props.markProps, undefined);
  // Typing into an empty paragraph whose mark is a tracked insertion: the text is bold, not "inserted".
  assert.ok(paragraphAt(doc, 3).props.markProps!.inserted);
  run(h, insertText(doc, pos(3, 0), 'x'));
  assert.deepEqual((paragraphAt(doc, 3).children[0] as Run).props, { bold: true });
  assert.ok(paragraphAt(doc, 3).props.markProps!.inserted, 'the mark keeps what it was');
  run(h, insertText(doc, pos(4, 0), 'y'));
  assert.deepEqual((paragraphAt(doc, 4).children[0] as Run).props, {});
  assert.deepEqual(shapeAt(doc, 4), [':y']);
});

test('Delete leaves the caret at the offset it had', async () => {
  const doc = await open(p('one two three'));
  const h = history(doc);
  assert.deepEqual(run(h, deleteForward(doc, pos(0, 4))).focus, pos(0, 4));
  assert.equal(texts(doc)[0], 'one wo three');
  assert.deepEqual(run(h, deleteForward(doc, pos(0, 4), 'word')).focus, pos(0, 4));
  assert.equal(texts(doc)[0], 'one three');
  assert.deepEqual(run(h, deleteForward(doc, pos(0, 0), 'word')).focus, pos(0, 0));
  assert.equal(texts(doc)[0], 'three');
});
