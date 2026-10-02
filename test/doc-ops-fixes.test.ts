// Defects found by a review of the document editing commands, one test (or a
// few) per defect: wrappers around blocks, what a deletion must leave behind,
// large tables, what Backspace takes, the formatting typed text takes over,
// what the commands refuse, and positions before a picture.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { paragraphText, type Block, type DocModel, type Paragraph, type Run, type Table, type TableRow, type Wrapper } from '../src/model/doc/types';
import {
  DocHistory,
  comparePositions,
  deleteBackward,
  deleteColumn,
  deleteForward,
  deleteRange,
  deleteRow,
  expandRangeForTables,
  graphemeBreaks,
  indentList,
  insertHyperlink,
  insertImage,
  insertPageBreak,
  insertText,
  movePosition,
  nextCharOffset,
  nextWordOffset,
  previousCharOffset,
  previousWordOffset,
  samePosition,
  setParaFormat,
  setRunFormat,
  snapOffset,
  splitParagraph,
  storyEnd,
  storyStart,
  toggleList,
  toggleRunFormat,
  wordAt,
  type DocEdit,
  type DocPosition,
  type DocRange,
} from '../src/ops/doc/index';
import { draftCounters } from '../src/ops/doc/draft';
import { dice, docPicture, open, p, paragraphAt, pos, r, range, run, shapeAt, tbl, tc, texts, tr } from './doc-ops-fixtures';

const history = (doc: DocModel) => new DocHistory<DocRange | undefined>(doc);
const before = (path: number[] | number, offset: number, story = 'body'): DocPosition => ({ ...pos(path, offset, story), before: true });

function rootOf(block: Block): Wrapper | undefined {
  let root = block.wrap;
  while (root && root.parent) root = root.parent;
  return root;
}

const sdt = (inner: string, id: number) => `<w:sdt><w:sdtPr><w:id w:val="${id}"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

// ------------------------------------------------------------------ 1 wrappers

test('a command over several paragraphs of one content control leaves them in one control', async () => {
  const commands: [string, (doc: DocModel, at: DocRange) => DocEdit][] = [
    ['setRunFormat', (doc, at) => setRunFormat(doc, at, { bold: true })],
    ['setParaFormat', (doc, at) => setParaFormat(doc, at, { align: 'center' })],
    ['toggleList', (doc, at) => toggleList(doc, at, 'bullet')],
    ['toggleRunFormat on', (doc, at) => toggleRunFormat(doc, at, 'italic')],
    ['indentList', (doc, at) => indentList(doc, at, 1)],
  ];
  for (const [name, command] of commands) {
    const doc = await open(p('zero') + sdt(p('one') + p('two') + p('three'), 1) + p('end'));
    const original = doc.body.blocks.slice();
    const h = history(doc);
    run(h, command(doc, range(pos(1, 1), pos(3, 2))));
    const [, a, b, c] = doc.body.blocks;
    assert.ok(a.wrap && a.wrap === b.wrap && b.wrap === c.wrap, `${name}: the three paragraphs share one wrapper`);
    assert.equal(a.wrap!.srcStart, undefined, name);
    assert.notEqual(a, original[1], name);
    h.undo();
    doc.body.blocks.forEach((block, i) => assert.equal(block, original[i], `${name}: undo puts block ${i} back`));
  }
});

test('turning a property off over a content control keeps the control whole too', async () => {
  const doc = await open(p('zero') + sdt(p(r('one', '<w:b/>')) + p(r('two', '<w:b/>')) + p(r('three', '<w:b/>')), 1) + p('end'));
  const h = history(doc);
  run(h, toggleRunFormat(doc, range(pos(1, 0), pos(3, 5)), 'bold'));
  const [, a, b, c] = doc.body.blocks;
  assert.deepEqual([shapeAt(doc, 1), shapeAt(doc, 2), shapeAt(doc, 3)], [[':one'], [':two'], [':three']]);
  assert.ok(a.wrap && a.wrap === b.wrap && b.wrap === c.wrap);
});

test('content controls inside each other: one copy of each, the inner one still inside the outer', async () => {
  const doc = await open(p('zero') + sdt(p('a') + sdt(p('b') + p('c'), 2) + p('d'), 1) + p('end'));
  const [, a0, b0, c0, d0] = doc.body.blocks;
  assert.ok(a0.wrap && b0.wrap && b0.wrap === c0.wrap && b0.wrap.parent === a0.wrap && d0.wrap === a0.wrap, 'as read');
  const h = history(doc);
  run(h, setRunFormat(doc, range(pos(1, 0), pos(4, 1)), { italic: true }));
  const [, a, b, c, d] = doc.body.blocks;
  assert.ok(a.wrap && a.wrap !== a0.wrap && a.wrap.srcStart === undefined);
  assert.equal(d.wrap, a.wrap);
  assert.equal(b.wrap, c.wrap);
  assert.ok(b.wrap && b.wrap !== a.wrap && b.wrap.parent === a.wrap);
  assert.equal(new Set([a, b, c, d].map(rootOf)).size, 1);
});

// ------------------------------------------------------------- 2 what stays

const MARKED_CELL = tc(
  p('<w:bookmarkStart w:id="5" w:name="bm"/>' + r('keepXgone') + '<w:commentRangeStart w:id="3"/>' + r('more') + '<w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r><w:strange/>') +
    '<!--between-->' +
    p(r('second') + '<w:bookmarkEnd w:id="5"/>'),
);

test('a cell emptied by a deletion keeps its bookmarks, comment marks, kept XML and comments', async () => {
  const doc = await open(tbl(tr(MARKED_CELL + tc('cell2')), [1000, 1000]) + p('tail text'));
  const h = history(doc);
  const after = run(h, deleteRange(doc, range(pos([0, 0, 0, 0], 5), pos(1, 4))));
  assert.deepEqual(texts(doc), ['', '', ' text']);
  const cell = (doc.body.blocks[0] as Table).rows[0].cells[0];
  assert.deepEqual(cell.blocks.map((block) => block.type), ['misc', 'paragraph']);
  assert.deepEqual(shapeAt(doc, [0, 0, 0, 1]), ['<bookmarkStart>', '<commentRangeStart>', '<commentRangeEnd>', ':<commentReference>', '<unknown>', '<bookmarkEnd>']);
  assert.deepEqual(after.focus, pos([0, 0, 0, 1], 0));
  h.undo();
  assert.deepEqual(texts(doc), ['keepXgonemore', 'second', 'cell2', 'tail text']);
});

test('cells emptied between two cells of a table keep what is not content', async () => {
  const doc = await open(tbl(tr(MARKED_CELL + tc('cell2')), [1000, 1000]) + p('tail'));
  run(history(doc), deleteRange(doc, range(pos([0, 0, 0, 0], 1), pos([0, 0, 1, 0], 1))));
  assert.deepEqual(texts(doc), ['', '', 'tail']);
  assert.deepEqual(shapeAt(doc, [0, 0, 0, 1]), ['<bookmarkStart>', '<commentRangeStart>', '<commentRangeEnd>', ':<commentReference>', '<unknown>', '<bookmarkEnd>']);
});

test('a table inside an emptied cell, and a table between the ends of a range, leave their bookmarks behind', async () => {
  const inner = tbl(tr(tc(p('<w:bookmarkStart w:id="8" w:name="deep"/>' + r('deep') + '<w:bookmarkEnd w:id="8"/>'))), [500]);
  const doc = await open(p('before') + tbl(tr(tc(inner + p('x')) + tc('y')), [1000, 1000]) + p('between') + tbl(tr(tc(p('<w:bookmarkStart w:id="9" w:name="mid"/>' + r('mid')))), [900]) + p(r('after') + '<w:bookmarkEnd w:id="9"/>'));
  const h = history(doc);
  run(h, deleteRange(doc, range(pos(0, 3), pos([1, 0, 0, 1], 1))));
  assert.deepEqual(shapeAt(doc, [1, 0, 0, 0]), ['<bookmarkStart>', '<bookmarkEnd>']);
  // From "between" over the second table into "after": the table goes, its bookmark start stays.
  run(h, deleteRange(doc, range(pos(2, 3), pos(4, 2))));
  assert.deepEqual(shapeAt(doc, 2), [':bet', '<bookmarkStart>', ':ter', '<bookmarkEnd>']);
});

const TABLE = tbl(tr(tc('a1') + tc('b1') + tc('c1')) + tr(tc('a2') + tc('b2') + tc('c2')), [1000, 1000, 1000]);

test('expandRangeForTables gives the range a deletion really covers', async () => {
  const doc = await open(p('before') + TABLE + p('after'));
  const same = range(pos(0, 1), pos(2, 3));
  assert.deepEqual(expandRangeForTables(doc, same), same);
  assert.deepEqual(expandRangeForTables(doc, range(pos([1, 0, 1, 0], 1), pos([1, 0, 1, 0], 2))), range(pos([1, 0, 1, 0], 1), pos([1, 0, 1, 0], 2)));
  // Out of a table: from the start of the cell the range begins in.
  assert.deepEqual(expandRangeForTables(doc, range(pos([1, 1, 1, 0], 1), pos(2, 2))), range(pos([1, 1, 1, 0], 0), pos(2, 2)));
  // Into a table: to the end of the cell it ends in; the direction of the range stays.
  assert.deepEqual(expandRangeForTables(doc, range(pos([1, 0, 1, 0], 1), pos(0, 3))), range(pos([1, 0, 1, 0], 2), pos(0, 3)));
  // Between cells: whole cells.
  assert.deepEqual(expandRangeForTables(doc, range(pos([1, 0, 1, 0], 1), pos([1, 1, 1, 0], 1))), range(pos([1, 0, 1, 0], 0), pos([1, 1, 1, 0], 2)));
  assert.deepEqual(expandRangeForTables(doc, range(pos([1, 1, 0, 0], 1), pos([1, 1, 2, 0], 1))), range(pos([1, 1, 0, 0], 0), pos([1, 1, 2, 0], 2)));
  // deleteRange does the same: text of the cell the range did not cover goes with it.
  const kept = await open(tbl(tr(tc('keepXgone') + tc('cell2')), [1000, 1000]) + p('tail text'));
  const at = range(pos([0, 0, 0, 0], 5), pos(1, 4));
  assert.deepEqual(expandRangeForTables(kept, at), range(pos([0, 0, 0, 0], 0), pos(1, 4)));
  run(history(kept), deleteRange(kept, at));
  assert.deepEqual(texts(kept), ['', '', ' text']);
  assert.throws(() => expandRangeForTables(doc, range(pos(0, 0), pos(9, 0))), RangeError);
});

test('rows, columns and tables that are deleted leave their bookmarks and comment marks in what stays', async () => {
  const marked = (text: string, id: number) => tc(p(`<w:bookmarkStart w:id="${id}" w:name="m${id}"/>` + r(text) + `<w:commentRangeStart w:id="${id}"/>`));
  const doc = await open(p('before') + tbl(tr(marked('a1', 1) + tc('b1')) + tr(marked('a2', 2) + tc('b2')), [1000, 1000]) + p(r('after') + '<w:bookmarkEnd w:id="1"/><w:bookmarkEnd w:id="2"/>'));
  const h = history(doc);
  run(h, deleteRow(doc, pos([1, 0, 0, 0], 0)));
  assert.deepEqual(shapeAt(doc, [1, 0, 0, 0]), ['<bookmarkStart>', '<commentRangeStart>', '<bookmarkStart>', ':a2', '<commentRangeStart>']);
  run(h, deleteColumn(doc, pos([1, 0, 0, 0], 0)));
  assert.deepEqual(shapeAt(doc, [1, 0, 0, 0]), ['<bookmarkStart>', '<commentRangeStart>', '<bookmarkStart>', '<commentRangeStart>', ':b2']);
  // The last row takes the table with it: the marks go into the paragraph after it.
  run(h, deleteRow(doc, pos([1, 0, 0, 0], 0)));
  assert.deepEqual(doc.body.blocks.map((block) => block.type), ['paragraph', 'paragraph']);
  assert.deepEqual(shapeAt(doc, 1), ['<bookmarkStart>', '<commentRangeStart>', '<bookmarkStart>', '<commentRangeStart>', ':after', '<bookmarkEnd>', '<bookmarkEnd>']);
  while (h.undo());
  assert.deepEqual(texts(doc), ['before', 'a1', 'b1', 'a2', 'b2', 'after']);
});

// ------------------------------------------------------------ 3 large tables

function bigTable(rows: number, columns: number): Table {
  const props = {};
  const cellProps = {};
  const list: TableRow[] = [];
  for (let i = 0; i < rows; i++) {
    const cells = [];
    for (let c = 0; c < columns; c++) {
      const run: Run = { type: 'run', props, content: [{ type: 'text', text: 'cell' }] };
      cells.push({ props: cellProps, blocks: [{ type: 'paragraph', props, children: [run] } as Paragraph] });
    }
    list.push({ props: {}, cells });
  }
  const grid: number[] = [];
  for (let c = 0; c < columns; c++) grid.push(1000);
  return { type: 'table', props: {}, grid, rows: list };
}

test('formatting every cell of a large table: one copy of the table, one operation, in little time', async (t) => {
  const doc = await open(p('x'));
  const ROWS = 10000;
  const table = bigTable(ROWS, 4);
  doc.body.blocks.splice(0, 0, table);
  const all = range(storyStart(doc, 'body')!, pos([0, ROWS - 1, 3, 0], 4));
  const counted = { ...draftCounters };
  const started = Date.now();
  const edit = setRunFormat(doc, all, { bold: true });
  const built = Date.now() - started;
  assert.equal(edit.ops.length, 1);
  assert.equal(draftCounters.tables - counted.tables, 1, 'the list of rows was copied once');
  assert.equal(draftCounters.rows - counted.rows, ROWS);
  assert.equal(draftCounters.cells - counted.cells, ROWS * 4);
  assert.equal(doc.body.blocks[0], table, 'building left the document alone');
  const h = history(doc);
  const applying = Date.now();
  h.applyEdit(edit, undefined);
  const applied = Date.now() - applying;
  const now = doc.body.blocks[0] as Table;
  assert.notEqual(now, table);
  for (const at of [0, 4999, ROWS - 1]) assert.deepEqual(shapeAt(doc, [0, at, 3, 0]), ['b:cell']);
  assert.equal(((table.rows[0].cells[0].blocks[0] as Paragraph).children[0] as Run).props.bold, undefined, 'the old table was not changed');
  h.undo();
  assert.equal(doc.body.blocks[0], table);
  t.diagnostic(`${ROWS} rows by 4: built in ${built} ms, applied in ${applied} ms`);
  // Generous: the quadratic version took seconds.
  assert.ok(built < 3000, `building took ${built} ms`);
});

test('one cell of a large table edited: the step copies one row, and typing on does not copy the table again per paragraph', async () => {
  const doc = await open(p('x'));
  const table = bigTable(2000, 3);
  doc.body.blocks.splice(0, 0, table);
  const counted = { ...draftCounters };
  const edit = insertText(doc, pos([0, 7, 1, 0], 2), 'one\ntwo\nthree');
  assert.equal(edit.ops.length, 1);
  assert.deepEqual([draftCounters.tables - counted.tables, draftCounters.rows - counted.rows, draftCounters.cells - counted.cells], [1, 1, 1]);
  const h = history(doc);
  run(h, edit);
  const now = doc.body.blocks[0] as Table;
  assert.deepEqual(now.rows[7].cells[1].blocks.map((block) => paragraphText(block as Paragraph)), ['ceone', 'two', 'threell']);
  // Everything beside the cell is shared with the table as it was.
  assert.equal(now.rows[6], table.rows[6]);
  assert.equal(now.rows[7].cells[0], table.rows[7].cells[0]);
  assert.notEqual(now.rows[7], table.rows[7]);
  h.undo();
  assert.equal(doc.body.blocks[0], table);
});

test('a nested table is edited through one copy of each table on the way', async () => {
  const inner = tbl(tr(tc('n1') + tc('n2')) + tr(tc('n3') + tc('n4')), [500, 500]);
  const doc = await open(p('zero') + tbl(tr(tc('a1') + tc(p('b1') + inner + p('b1 end'))) + tr(tc('a2') + tc('b2')), [1000, 1000]) + p('last'));
  const original = doc.body.blocks[1] as Table;
  const h = history(doc);
  const edit = setRunFormat(doc, range(pos([1, 0, 0, 0], 0), pos([1, 1, 1, 0], 2)), { bold: true });
  assert.equal(edit.ops.length, 1);
  run(h, edit);
  assert.deepEqual(shapeAt(doc, [1, 0, 1, 1, 1, 0, 0]), ['b:n3']);
  assert.deepEqual(shapeAt(doc, [1, 1, 1, 0]), ['b:b2']);
  assert.deepEqual(shapeAt(doc, [1, 0, 1, 2]), ['b:b1 end']);
  assert.equal(((original.rows[0].cells[1].blocks[1] as Table).rows[1].cells[0].blocks[0] as Paragraph).children.length, 1);
  assert.equal((((original.rows[0].cells[1].blocks[1] as Table).rows[1].cells[0].blocks[0] as Paragraph).children[0] as Run).props.bold, undefined);
  h.undo();
  assert.equal(doc.body.blocks[1], original);
});

// ------------------------------------------------------------- 4 Backspace

const LRPB = '<w:r><w:lastRenderedPageBreak/><w:t>def</w:t></w:r>';

test('Backspace and Delete take what shows, not a layout hint beside the caret', async () => {
  const doc = await open(p(r('abc') + LRPB));
  const h = history(doc);
  assert.equal(paragraphText(paragraphAt(doc, 0)), 'abcdef');
  const after = run(h, deleteBackward(doc, pos(0, 3)));
  assert.deepEqual(shapeAt(doc, 0), [':abdef'], 'the "c" goes, and the stale hint with it');
  assert.deepEqual(after.focus, pos(0, 2));
  h.undo();
  run(h, deleteBackward(doc, pos(0, 3), 'word'));
  assert.deepEqual(shapeAt(doc, 0), [':def']);
  h.undo();
  // A caret that stands before a bookmark: Delete takes the "d" behind it, keeps the bookmark, drops the hint.
  const marked = await open(p(r('abc') + '<w:bookmarkStart w:id="1" w:name="b"/>' + LRPB + '<w:bookmarkEnd w:id="1"/>'));
  run(history(marked), deleteForward(marked, pos(0, 3)));
  assert.deepEqual(shapeAt(marked, 0), [':abc', '<bookmarkStart>', ':ef', '<bookmarkEnd>']);
});

test('Backspace passes over spelling marks (dropped), bookmarks (kept) and a soft hyphen (kept)', async () => {
  const spelled = await open(p(r('abc') + '<w:proofErr w:type="spellStart"/>' + r('def') + '<w:proofErr w:type="spellEnd"/>'));
  run(history(spelled), deleteBackward(spelled, pos(0, 3)));
  assert.deepEqual(shapeAt(spelled, 0), [':abdef', '<proofError>']);

  const marked = await open(p('<w:bookmarkStart w:id="1" w:name="b"/>' + r('abc') + '<w:bookmarkEnd w:id="1"/>' + r('def')));
  run(history(marked), deleteBackward(marked, pos(0, 3)));
  assert.deepEqual(shapeAt(marked, 0), ['<bookmarkStart>', ':ab', '<bookmarkEnd>', ':def']);

  const soft = await open(p('<w:r><w:t>abc</w:t><w:softHyphen/><w:t>def</w:t></w:r>'));
  const h = history(soft);
  run(h, deleteBackward(soft, pos(0, 3)));
  assert.deepEqual(shapeAt(soft, 0), [':ab<softHyphen>def']);
  h.undo();
  // A range over it takes it.
  run(h, deleteRange(soft, range(pos(0, 2), pos(0, 4))));
  assert.deepEqual(shapeAt(soft, 0), [':abef']);

  const hard = await open(p('<w:r><w:t>abc</w:t><w:noBreakHyphen/><w:t>def</w:t></w:r>'));
  run(history(hard), deleteBackward(hard, pos(0, 4)));
  assert.deepEqual(shapeAt(hard, 0), [':abcdef']);
});

// -------------------------------------------------- 5 formatting of typed text

const field = (code: string, result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

test('text typed after a hyperlink, a field or a content control does not take their formatting', async () => {
  const link = (text: string, rPr: string) => `<w:hyperlink r:id="rId9">${r(text, rPr)}</w:hyperlink>`;
  const doc = await open({
    body:
      p(r('plain ', '<w:sz w:val="32"/>') + link('link', '<w:rStyle w:val="Strong"/><w:color w:val="0563C1"/><w:u w:val="single"/>')) +
      p(link('only', '<w:u w:val="single"/>'), '<w:rPr><w:i/></w:rPr>') +
      p(`<w:sdt><w:sdtPr><w:alias w:val="c"/></w:sdtPr><w:sdtContent>${r('ctl', '<w:b/>')}</w:sdtContent></w:sdt>`) +
      p(r('see ') + field('REF x', r('res', '<w:b/>'))),
    rels: [['rId9', 'hyperlink', 'https://example.com/', true]],
  });
  const h = history(doc);
  run(h, insertText(doc, pos(0, 10), 'X'));
  assert.deepEqual(shapeAt(doc, 0), ['32:plain ', 'u{Strong}:link@hyperlink', '32:X']);
  // Inside the link the text stays in it and is formatted like it.
  run(h, insertText(doc, pos(0, 8), 'Y'));
  assert.deepEqual(shapeAt(doc, 0), ['32:plain ', 'u{Strong}:liYnk@hyperlink', '32:X']);
  // Nothing outside the link to the left: the formatting of the paragraph mark, before the link as after it.
  run(h, insertText(doc, pos(1, 4), 'A'));
  run(h, insertText(doc, pos(1, 0), 'B'));
  assert.deepEqual(shapeAt(doc, 1), ['i:B', 'u:only@hyperlink', 'i:A']);
  run(h, insertText(doc, pos(2, 3), 'C'));
  assert.deepEqual(shapeAt(doc, 2), ['b:ctl@sdt', ':C']);
  run(h, insertText(doc, pos(3, 7), 'D'));
  assert.deepEqual(shapeAt(doc, 3).slice(-2), [':<end>', ':D']);
});

// ------------------------------------------------------------- 6 pictures

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4]);
const JPEG = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0, 16, 0x4a, 0x46, 0x49, 0x46]);
const GIF = new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 1, 0, 1, 0]);
const BMP = new Uint8Array([0x42, 0x4d, 30, 0, 0, 0, 0, 0, 0, 0]);
const WEBP = new Uint8Array([0x52, 0x49, 0x46, 0x46, 4, 0, 0, 0, 0x57, 0x45, 0x42, 0x50, 0x56, 0x50]);

test('insertImage: only pictures a browser shows, whose bytes are what the type says, of a sane size', async () => {
  const doc = await open(p('text'));
  const picture = docPicture(doc);
  const refused = (image: unknown) => {
    assert.throws(() => insertImage(doc, pos(0, 0), image as never), RangeError, JSON.stringify(image));
    assert.equal(docPicture(doc), picture);
  };
  for (const type of ['toString', 'constructor', '__proto__', 'hasOwnProperty', 'image/tiff', 'image/x-emf', 'image/svg+xml', 'IMAGE/PNG', '']) refused({ bytes: PNG, type, width: 100, height: 100 });
  refused({ bytes: JPEG, type: 'image/png', width: 100, height: 100 });
  refused({ bytes: PNG, type: 'image/jpeg', width: 100, height: 100 });
  refused({ bytes: PNG.slice(0, 4), type: 'image/png', width: 100, height: 100 });
  refused({ bytes: new Uint8Array([0x52, 0x49, 0x46, 0x46, 4, 0, 0, 0, 0x57, 0x41, 0x56, 0x45]), type: 'image/webp', width: 100, height: 100 });
  for (const size of [0, -1, 1.5, NaN, Infinity, 1e12, '100']) {
    refused({ bytes: PNG, type: 'image/png', width: size, height: 100 });
    refused({ bytes: PNG, type: 'image/png', width: 100, height: size });
  }
  refused({ bytes: PNG, type: 'image/png', width: 100, height: 100, description: 5 });
  refused(null);
  const h = history(doc);
  const all: [Uint8Array, string, string][] = [
    [PNG, 'image/png', 'png'],
    [JPEG, 'image/jpeg', 'jpeg'],
    [GIF, 'image/gif', 'gif'],
    [BMP, 'image/bmp', 'bmp'],
    [WEBP, 'image/webp', 'webp'],
  ];
  all.forEach(([bytes, type, extension], i) => {
    run(h, insertImage(doc, pos(0, 0), { bytes, type, width: 9525, height: 9525 }));
    assert.ok(doc.media!.has(`/word/media/image${i + 1}.${extension}`), type);
  });
});

test('insertImage: the id of the drawing is beyond every id in the document, wherever it is', async () => {
  const drawing = (id: number, inner = '') => `<w:r><w:drawing><wp:inline><wp:extent cx="95250" cy="95250"/><wp:docPr id="${id}" name="d"/>${inner}</wp:inline></w:drawing></w:r>`;
  const nested = '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="90" name="p"/></pic:nvPicPr></pic:pic></a:graphicData></a:graphic>';
  const vml = (inner: string) => `<w:r><w:pict><v:shape><v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`;
  const ids = async (spec: Parameters<typeof open>[0]): Promise<string[]> => {
    const doc = await open(spec);
    const h = history(doc);
    run(h, insertImage(doc, pos(0, 0), { bytes: PNG, type: 'image/png', width: 9525, height: 9525 }));
    const paragraph = paragraphAt(doc, 0);
    for (const inline of paragraph.children) {
      if (inline.type !== 'run') continue;
      for (const item of inline.content) {
        if (item.type !== 'drawing' || item.name === 'd') continue;
        const text = JSON.stringify(item.xml, (_key, value) => (value instanceof Map ? Array.from(value.entries()) : value));
        return (text.match(/\["id","(\d+)"\]/g) || []).map((found) => found.replace(/\D/g, ''));
      }
    }
    return [];
  };
  assert.deepEqual(await ids(p(r('x') + drawing(7))), ['8', '8']);
  assert.deepEqual(await ids(p(r('x') + drawing(7, nested))), ['91', '91']);
  assert.deepEqual(await ids(p(r('x') + vml(p(drawing(60))))), ['61', '61']);
  assert.deepEqual(await ids({ body: p('x') + '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="9000" w:h="9000"/></w:sectPr>', headers: { rIdH: p(drawing(120)) } }), ['121', '121']);
  assert.deepEqual(await ids({ body: p(r('x') + '<w:r><w:footnoteReference w:id="2"/></w:r>'), footnotes: `<w:footnote w:id="2">${p(drawing(130))}</w:footnote>` }), ['131', '131']);
  assert.deepEqual(await ids({ body: p('x'), comments: `<w:comment w:id="1" w:author="A">${p(drawing(140))}</w:comment>` }), ['141', '141']);
  const alternate = `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline><wp:docPr id="300" name="d"/></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>`;
  assert.deepEqual(await ids(p(r('x') + alternate)), ['301', '301']);
});

test('insertImage: the name of the new part is not one the document or the package has', async () => {
  const doc = await open({ body: p('x') + '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="9000" w:h="9000"/></w:sectPr>', headers: { rIdH: p('head') } });
  const header = doc.headers.get('rIdH')!;
  header.relationships = new Map(header.relationships);
  header.relationships.set('rId77', { id: 'rId77', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', target: 'media/image12.jpeg', external: false, part: '/word/media/image12.jpeg' });
  const h = history(doc);
  run(h, insertImage(doc, pos(0, 0), { bytes: PNG, type: 'image/png', width: 9525, height: 9525 }));
  assert.deepEqual(Array.from(doc.media!.keys()), ['/word/media/image13.png']);
  const asked: string[] = [];
  const taken = (part: string) => {
    asked.push(part);
    return part === '/word/media/image14.png' || part === '/word/media/image15.png';
  };
  run(h, insertImage(doc, pos(0, 0), { bytes: PNG, type: 'image/png', width: 9525, height: 9525 }, { taken }));
  assert.deepEqual(Array.from(doc.media!.keys()), ['/word/media/image13.png', '/word/media/image16.png']);
  assert.deepEqual(asked, ['/word/media/image14.png', '/word/media/image15.png', '/word/media/image16.png']);
  // The relationship does not take an id that is in use either.
  assert.equal(new Set(Array.from(doc.body.relationships.keys())).size, doc.body.relationships.size);
});

// ---------------------------------------------------------------- 7 patches

test('a patch of character or paragraph formatting is checked field by field', async () => {
  const doc = await open(p('some text') + p('more'));
  const picture = docPicture(doc);
  const at = range(pos(0, 1), pos(1, 2));
  const refusedRun = (patch: unknown) => {
    assert.throws(() => setRunFormat(doc, at, patch as never), RangeError, 'run ' + JSON.stringify(patch));
    if (patch !== null) assert.throws(() => insertText(doc, pos(0, 1), 'x', patch as never), RangeError, 'typing ' + JSON.stringify(patch));
    assert.equal(docPicture(doc), picture);
  };
  const refusedPara = (patch: unknown) => {
    assert.throws(() => setParaFormat(doc, at, patch as never), RangeError, 'paragraph ' + JSON.stringify(patch));
    assert.equal(docPicture(doc), picture);
  };
  for (const patch of [
    { bold: 'yes' },
    { italic: 1 },
    { strike: {} },
    { color: 5 },
    { color: 'red' },
    { color: {} },
    { color: { rgb: 'XYZ' } },
    { color: { theme: 5 } },
    { size: 1.5 },
    { size: 0 },
    { size: '12' },
    { size: 100000 },
    { style: 5 },
    { style: 'Nope' },
    { style: 'Heading1' },
    { highlight: 'x1' },
    { highlight: 'mauve' },
    { underline: 5 },
    { underline: 'wobbly' },
    { vertAlign: 'up' },
    { font: 5 },
    { bogus: 1 },
    'bold',
    null,
  ]) {
    refusedRun(patch);
  }
  for (const patch of [
    { align: {} },
    { align: 'middle' },
    { align: 5 },
    { style: 5 },
    { style: 'Nope' },
    { style: 'Strong' },
    { indent: { left: 'x' } },
    { indent: { left: 1.5 } },
    { indent: { left: 1e9 } },
    { indent: { bogus: 1 } },
    { indent: 5 },
    { spacing: { bogus: 2 } },
    { spacing: { before: -5 } },
    { spacing: { lineRule: 'loose' } },
    { spacing: { beforeAuto: 1 } },
    { numbering: { numId: 999 } },
    { numbering: { numId: 1, level: 9 } },
    { numbering: { numId: 1, bogus: 2 } },
    { numbering: 5 },
    { keepNext: 'yes' },
    { outlineLevel: 12 },
    { bogus: true },
    null,
  ]) {
    refusedPara(patch);
  }
  // What is valid still goes through.
  const h = history(doc);
  run(h, setRunFormat(doc, at, { color: { theme: 'accent1' }, underline: 'double', highlight: 'yellow', style: 'Strong', size: 24, bold: null }));
  run(h, setRunFormat(doc, at, { color: 'auto', underline: false, highlight: null, style: null }));
  run(h, setParaFormat(doc, at, { align: 'both', style: 'Heading1', indent: { left: -360, hanging: 360 }, spacing: { before: 120, line: 240, lineRule: 'auto' }, numbering: { numId: 0 }, keepNext: true, outlineLevel: 0 }));
  assert.deepEqual(paragraphAt(doc, 0).props.indent, { left: -360, hanging: 360 });
});

test('a missing position or range, and an unknown unit or direction, are refused with a RangeError', async () => {
  const doc = await open(p('some text'));
  const at = pos(0, 1);
  const calls: [string, () => unknown][] = [
    ['deleteRange null', () => deleteRange(doc, null as never)],
    ['deleteRange undefined', () => deleteRange(doc, undefined as never)],
    ['deleteRange half', () => deleteRange(doc, { anchor: at, focus: null } as never)],
    ['insertText null', () => insertText(doc, null as never, 'x')],
    ['insertText half', () => insertText(doc, { anchor: undefined, focus: at } as never, 'x')],
    ['setRunFormat null', () => setRunFormat(doc, null as never, { bold: true })],
    ['setRunFormat half', () => setRunFormat(doc, { anchor: at, focus: undefined } as never, { bold: true })],
    ['toggleRunFormat null', () => toggleRunFormat(doc, null as never, 'bold')],
    ['setParaFormat null', () => setParaFormat(doc, null as never, { align: 'center' })],
    ['toggleList null', () => toggleList(doc, undefined as never, 'bullet')],
    ['indentList null', () => indentList(doc, null as never, 1)],
    ['insertHyperlink null', () => insertHyperlink(doc, null as never, 'https://example.com/')],
    ['deleteBackward null', () => deleteBackward(doc, null as never)],
    ['splitParagraph null', () => splitParagraph(doc, undefined as never)],
    ['comparePositions null', () => comparePositions(null as never, at)],
    ['comparePositions no path', () => comparePositions(at, { story: 'body', offset: 0 } as never)],
    ['movePosition unit', () => movePosition(doc, at, 'sentence' as never, 1)],
    ['movePosition direction', () => movePosition(doc, at, 'char', 2 as never)],
    ['movePosition direction 0', () => movePosition(doc, at, 'word', 0 as never)],
    ['deleteBackward unit', () => deleteBackward(doc, at, 'line' as never)],
    ['deleteForward unit', () => deleteForward(doc, at, 'paragraph' as never)],
    ['expandRangeForTables null', () => expandRangeForTables(doc, null as never)],
  ];
  for (const [name, call] of calls) assert.throws(call, RangeError, name);
});

// -------------------------------------------------------------- 8 hyperlinks

test('insertHyperlink refuses a target that is not safe to follow', async () => {
  const doc = await open(p('some text'));
  const picture = docPicture(doc);
  const at = range(pos(0, 0), pos(0, 4));
  for (const target of ['javascript:alert(1)', ' JavaScript:alert(1)', 'java\tscript:alert(1)', 'data:text/html,x', 'file:///etc/passwd', 'vbscript:x', 'relative/page', 'https:nohost', '//example.com/', '#', '']) {
    assert.throws(() => insertHyperlink(doc, at, target), RangeError, target);
    assert.equal(docPicture(doc), picture);
  }
  const h = history(doc);
  run(h, insertHyperlink(doc, at, ' https://example.com/a '));
  assert.deepEqual(Array.from(doc.body.relationships.values()).filter((rel) => rel.external).map((rel) => rel.target), ['https://example.com/a']);
  run(h, insertHyperlink(doc, range(pos(0, 5), pos(0, 9)), '#mark'));
  const last = paragraphAt(doc, 0).children[paragraphAt(doc, 0).children.length - 1];
  assert.deepEqual([last.wrap!.kind, (last.wrap as { anchor?: string }).anchor, (last.wrap as { rId?: string }).rId], ['hyperlink', 'mark', undefined]);
  run(h, insertHyperlink(doc, range(pos(0, 5), pos(0, 9)), 'mailto:someone@example.com'));
  run(h, insertHyperlink(doc, range(pos(0, 5), pos(0, 9)), 'tel:+123456'));
});

// ---------------------------------------------------------- 9 characters, words

const stops = (text: string, step: (text: string, offset: number) => number, from: number, to: number): number[] => {
  const out = [from];
  let at = from;
  for (let guard = 0; guard <= text.length && at !== to; guard++) {
    const next = step(text, at);
    assert.ok(from < to ? next > at : next < at, `no progress in ${JSON.stringify(text)} at ${at}`);
    at = next;
    out.push(at);
  }
  return out;
};

test('the places a caret stops at are the same going forward, going back and when an offset is snapped', () => {
  const d = dice(20261003);
  const alphabet = ['a', 'b', 'Z', ' ', ' ', ',', '.', "'", '’', '\u0301', '\u0308', '\u200d', '\ufe0f', '\u200c', '😀', '👩', '👍', '\u{1f3fd}', '🇹', '🇷', '🇩', '\u1112', '\u1161', '\u11ab', '한', '가', '日', '本', 'カ', 'ん', '\ud83d', '\udc00', '\n', '-', '1', 'é', '\u20e3', '❤'];
  for (let n = 0; n < 20000; n++) {
    let text = '';
    const length = d.int(12);
    for (let i = 0; i < length; i++) text += d.pick(alphabet);
    const forward = stops(text, nextCharOffset, 0, text.length);
    const backward = stops(text, previousCharOffset, text.length, 0).reverse();
    assert.deepEqual(forward, backward, `characters of ${JSON.stringify(text)}`);
    assert.deepEqual(graphemeBreaks(text), forward, JSON.stringify(text));
    for (let offset = 0; offset <= text.length; offset++) {
      const snapped = snapOffset(text, offset);
      assert.ok(snapped >= offset && forward.indexOf(snapped) >= 0, `snapOffset(${JSON.stringify(text)}, ${offset}) = ${snapped}`);
      if (forward.indexOf(offset) >= 0) assert.equal(snapped, offset);
    }
    const wordsForward = stops(text, nextWordOffset, 0, text.length);
    const wordsBackward = stops(text, previousWordOffset, text.length, 0).reverse();
    assert.deepEqual(wordsForward, wordsBackward, `words of ${JSON.stringify(text)}`);
    for (const stop of wordsForward) assert.ok(forward.indexOf(stop) >= 0, `a word of ${JSON.stringify(text)} starts inside a character`);
    for (let offset = 0; offset < text.length; offset++) {
      const [start, end] = wordAt(text, offset);
      assert.ok(start <= offset && end > offset && forward.indexOf(start) >= 0 && forward.indexOf(end) >= 0, `wordAt(${JSON.stringify(text)}, ${offset})`);
    }
  }
});

test('flags, skin tones, variation selectors, joined emoji and Hangul jamo are one character each', () => {
  assert.deepEqual(graphemeBreaks('a😀b'), [0, 1, 3, 4]);
  assert.deepEqual(graphemeBreaks('🇹🇷🇩🇪'), [0, 4, 8]);
  assert.deepEqual(graphemeBreaks('🇹🇷🇩'), [0, 4, 6]);
  assert.deepEqual(graphemeBreaks('x👍\u{1f3fd}y'), [0, 1, 5, 6]);
  assert.deepEqual(graphemeBreaks('❤\ufe0f!'), [0, 2, 3]);
  assert.deepEqual(graphemeBreaks('👨\u200d👩\u200d👧.'), [0, 8, 9]);
  assert.deepEqual(graphemeBreaks('1\ufe0f\u20e32'), [0, 3, 4]);
  assert.deepEqual(graphemeBreaks('\u1112\u1161\u11ab\u1112\u1161'), [0, 3, 5]);
  assert.deepEqual(graphemeBreaks('가\u11a8나'), [0, 2, 3]);
  assert.deepEqual(graphemeBreaks('\r\n\n'), [0, 2, 3]);
  assert.equal(previousCharOffset('🇹🇷🇩🇪', 8), 4);
  assert.equal(snapOffset('🇹🇷🇩🇪', 5), 8);
  assert.equal(snapOffset('a\u200db', 2), 3, 'never back to before the offset');
  assert.equal(snapOffset('ab\u0301c', 2), 3);
});

test('an apostrophe between letters does not end a word; ideographs are a word each', () => {
  const text = 'İstanbul’da kaldı, don\'t \'quoted\'';
  assert.equal(nextWordOffset(text, 0), 12);
  assert.equal(previousWordOffset(text, 11), 0);
  assert.deepEqual(wordAt(text, 9), [0, 11]);
  assert.equal(nextWordOffset(text, 19), 25);
  assert.equal(nextWordOffset(text, 25), 26, 'an apostrophe that opens a quote is not part of the word');
  assert.equal(previousWordOffset(text, text.length), text.length - 1);
  assert.equal(previousWordOffset(text, text.length - 1), 26);
  assert.deepEqual(stops('日本語 かな。x', nextWordOffset, 0, 8), [0, 1, 2, 4, 5, 6, 7, 8]);
  assert.deepEqual(wordAt('日本語', 1), [1, 2]);
});

// ---------------------------------------------------------- 10 Word's behaviour

test('a range that covers a table completely removes it; the story keeps a paragraph', async () => {
  const doc = await open(p('one') + TABLE + p('two'));
  const h = history(doc);
  const after = run(h, deleteRange(doc, range(storyStart(doc, 'body')!, storyEnd(doc, 'body')!)));
  assert.deepEqual(doc.body.blocks.map((block) => block.type), ['paragraph']);
  assert.deepEqual(texts(doc), ['']);
  assert.deepEqual(after.focus, pos(0, 0));
  h.undo();

  // The story begins with the table.
  const first = await open(TABLE + p('end'));
  const h2 = history(first);
  const caret = run(h2, deleteRange(first, range(storyStart(first, 'body')!, storyEnd(first, 'body')!)));
  assert.deepEqual(first.body.blocks.map((block) => block.type), ['paragraph']);
  assert.deepEqual(texts(first), ['']);
  assert.deepEqual(caret.focus, pos(0, 0));
  h2.undo();
  // Not from its very start: the table stays, its cells are emptied.
  run(h2, deleteRange(first, range(pos([0, 0, 0, 0], 1), storyEnd(first, 'body')!)));
  assert.deepEqual(first.body.blocks.map((block) => block.type), ['table', 'paragraph']);
  assert.deepEqual(texts(first), ['', '', '', '', '', '', '']);
  h2.undo();
  // To the very end of a table from before it.
  const last = await open(p('start') + TABLE + p('end'));
  run(history(last), deleteRange(last, range(pos(0, 2), pos([1, 1, 2, 0], 2))));
  assert.deepEqual(last.body.blocks.map((block) => block.type), ['paragraph', 'paragraph']);
  assert.deepEqual(texts(last), ['st', 'end']);
  // From the start of one table to the end of another, with nothing but them: an empty paragraph stays.
  const two = await open(TABLE + p('mid') + TABLE + p(''));
  run(history(two), deleteRange(two, range(pos([0, 0, 0, 0], 0), pos([2, 1, 2, 0], 2))));
  assert.deepEqual(two.body.blocks.map((block) => block.type), ['paragraph', 'paragraph']);
  assert.deepEqual(texts(two), ['', '']);
});

test('Enter in the middle of a paragraph: both halves keep the style and the direct formatting', async () => {
  const doc = await open(p('Title', '<w:pStyle w:val="Heading1"/><w:jc w:val="center"/><w:ind w:left="720"/><w:spacing w:before="240"/>'));
  const h = history(doc);
  run(h, splitParagraph(doc, pos(0, 2)));
  for (const at of [0, 1]) assert.deepEqual(paragraphAt(doc, at).props, { style: 'Heading1', align: 'center', indent: { left: 720 }, spacing: { before: 240 } });
  // At the end the paragraph that follows takes the next style, and none of the direct formatting.
  run(h, splitParagraph(doc, pos(1, 3)));
  assert.deepEqual(paragraphAt(doc, 2).props, {});
});

test('a page break cannot be put into a table cell', async () => {
  const doc = await open(p('before') + TABLE + p('after'));
  const picture = docPicture(doc);
  assert.throws(() => insertPageBreak(doc, pos([1, 0, 0, 0], 1)), RangeError);
  assert.equal(docPicture(doc), picture);
  assert.equal(insertPageBreak(doc, pos(0, 1)).ops.length, 1);
});

test('character formatting at a caret: on the mark of an empty paragraph, nothing in a paragraph with text', async () => {
  const doc = await open(p('') + p('text'));
  const h = history(doc);
  assert.equal(setRunFormat(doc, range(pos(1, 2), pos(1, 2)), { bold: true }).ops.length, 0);
  run(h, setRunFormat(doc, range(pos(0, 0), pos(0, 0)), { bold: true, size: 40 }));
  assert.deepEqual(paragraphAt(doc, 0).props.markProps, { bold: true, boldCs: true, size: 40, sizeCs: 40 });
  run(h, toggleRunFormat(doc, range(pos(0, 0), pos(0, 0)), 'italic'));
  assert.equal(paragraphAt(doc, 0).props.markProps!.italic, true);
  run(h, toggleRunFormat(doc, range(pos(0, 0), pos(0, 0)), 'italic'));
  assert.equal(paragraphAt(doc, 0).props.markProps!.italic, undefined);
  run(h, insertText(doc, pos(0, 0), 'typed'));
  assert.deepEqual(shapeAt(doc, 0), ['b40:typed']);
});

const NOTES =
  '<w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  [2, 3, 4, 5].map((id) => `<w:footnote w:id="${id}"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> note ${id}</w:t></w:r></w:p></w:footnote>`).join('');
const ref = (id: number, kind = 'footnote') => `<w:r><w:${kind}Reference w:id="${id}"/></w:r>`;

test('a note whose reference is deleted leaves the document with it, and comes back with it', async () => {
  const doc = await open({
    body: p(r('ab') + ref(2) + r('cd') + ref(3)) + p(r('second') + ref(4) + ref(4)) + tbl(tr(tc(p(r('cell') + ref(5)))) + tr(tc('other')), [1000]) + p(r('end') + ref(1, 'endnote')),
    footnotes: NOTES,
    endnotes: '<w:endnote w:id="1"><w:p><w:r><w:t>end note</w:t></w:r></w:p></w:endnote>',
  });
  const h = history(doc);
  const items = doc.footnotes!.items;
  const notes = new Map(items);
  const ids = () => Array.from(doc.footnotes!.items.keys());
  assert.deepEqual(ids(), [0, 1, 2, 3, 4, 5]);
  run(h, deleteBackward(doc, pos(0, 2)));
  assert.deepEqual(ids(), [0, 1, 3, 4, 5]);
  assert.equal(doc.footnotes!.modified, true);
  h.undo();
  assert.equal(doc.footnotes!.items, items, 'undo puts the very collection back');
  assert.deepEqual(ids(), [0, 1, 2, 3, 4, 5]);
  assert.equal(doc.footnotes!.items.get(2), notes.get(2));
  assert.equal(doc.footnotes!.modified, undefined);
  h.redo();
  assert.deepEqual(ids(), [0, 1, 3, 4, 5]);
  h.undo();
  // A range, typing over, and one of two references to the same note (which then stays).
  run(h, insertText(doc, range(pos(0, 1), pos(0, 4)), 'X'));
  assert.deepEqual(ids(), [0, 1, 4, 5]);
  run(h, deleteBackward(doc, pos(1, 6)));
  assert.deepEqual(ids(), [0, 1, 4, 5]);
  run(h, deleteBackward(doc, pos(1, 6)));
  assert.deepEqual(ids(), [0, 1, 5]);
  // With the row of a table, and with a paragraph joined away.
  run(h, deleteRow(doc, pos([2, 0, 0, 0], 0)));
  assert.deepEqual(ids(), [0, 1]);
  assert.equal(doc.endnotes!.items.size, 1);
  run(h, deleteRange(doc, range(pos(1, 0), pos(3, 3))));
  assert.equal(doc.endnotes!.items.size, 0);
  assert.equal(doc.endnotes!.modified, true);
  while (h.undo());
  assert.deepEqual(ids(), [0, 1, 2, 3, 4, 5]);
  assert.equal(doc.footnotes!.items, items);
  notes.forEach((note, id) => assert.equal(doc.footnotes!.items.get(id), note));
  assert.equal(doc.endnotes!.items.size, 1);
  assert.equal(doc.endnotes!.modified, undefined);
});

// ------------------------------------------------ 11 before a picture, hidden text

const DRAWING = '<w:r><w:drawing><wp:inline><wp:extent cx="95250" cy="95250"/><wp:docPr id="1" name="d"/></wp:inline></w:drawing></w:r>';

test('a position can stand before what has no width at its offset', async () => {
  const doc = await open(p(DRAWING + r('cd')) + p(r('ab') + DRAWING + r('cd')));
  const h = history(doc);
  // Typing before a picture at the start of a paragraph; the caret stays before it.
  let caret = run(h, insertText(doc, before(0, 0), 'X'));
  assert.deepEqual(shapeAt(doc, 0), [':X', ':<drawing>', ':cd']);
  assert.deepEqual(caret.focus, before(0, 1));
  caret = run(h, insertText(doc, caret.focus, 'Y'));
  assert.deepEqual(shapeAt(doc, 0), [':XY', ':<drawing>', ':cd']);
  assert.deepEqual(caret.focus, before(0, 2));
  // Without the flag the position is after the picture, as it always was.
  caret = run(h, insertText(doc, pos(0, 2), 'Z'));
  assert.deepEqual(shapeAt(doc, 0), [':XY', ':<drawing>', ':Zcd']);
  assert.deepEqual(caret.focus, pos(0, 3));

  assert.ok(comparePositions(before(1, 2), pos(1, 2)) < 0);
  assert.ok(comparePositions(pos(1, 2), before(1, 2)) > 0);
  assert.equal(comparePositions(before(1, 2), before(1, 2)), 0);
  assert.ok(comparePositions(pos(1, 1), before(1, 2)) < 0);
  assert.equal(samePosition(before(1, 2), pos(1, 2)), false);
  assert.equal(samePosition(before(1, 2), before(1, 2)), true);

  // The picture alone as a range.
  run(h, setRunFormat(doc, range(before(1, 2), pos(1, 2)), { bold: true }));
  assert.deepEqual(shapeAt(doc, 1), [':ab', 'b:<drawing>', ':cd']);
  h.undo();
  run(h, setRunFormat(doc, range(before(1, 2), pos(1, 4)), { bold: true }));
  assert.deepEqual(shapeAt(doc, 1), [':ab', 'b:<drawing>', 'b:cd']);
  h.undo();
  run(h, setRunFormat(doc, range(pos(1, 2), pos(1, 4)), { bold: true }));
  assert.deepEqual(shapeAt(doc, 1), [':ab', ':<drawing>', 'b:cd']);
  h.undo();
  run(h, setRunFormat(doc, range(pos(1, 0), before(1, 2)), { bold: true }));
  assert.deepEqual(shapeAt(doc, 1), ['b:ab', ':<drawing>', ':cd']);
  h.undo();
  caret = run(h, deleteRange(doc, range(pos(1, 2), before(1, 2))));
  assert.deepEqual(shapeAt(doc, 1), [':abcd']);
  assert.deepEqual(caret.focus, pos(1, 2));
  h.undo();
  // With the flag on its end a range stops before the picture; with it on its start the range takes the picture in.
  run(h, deleteRange(doc, range(pos(1, 1), before(1, 2))));
  assert.deepEqual(shapeAt(doc, 1), [':a', ':<drawing>', ':cd']);
  h.undo();
  run(h, deleteRange(doc, range(before(1, 2), pos(1, 3))));
  assert.deepEqual(shapeAt(doc, 1), [':abd']);
  h.undo();

  caret = run(h, deleteForward(doc, before(1, 2)));
  assert.deepEqual(shapeAt(doc, 1), [':abcd']);
  h.undo();
  caret = run(h, deleteBackward(doc, before(1, 2)));
  assert.deepEqual(shapeAt(doc, 1), [':a', ':<drawing>', ':cd']);
  assert.deepEqual(caret.focus, before(1, 1));
  h.undo();
  caret = run(h, splitParagraph(doc, before(1, 2)));
  assert.deepEqual([shapeAt(doc, 1), shapeAt(doc, 2)], [[':ab'], [':<drawing>', ':cd']]);
  assert.deepEqual(caret.focus, before(2, 0));
  h.undo();
  caret = run(h, splitParagraph(doc, pos(1, 2)));
  assert.deepEqual([shapeAt(doc, 1), shapeAt(doc, 2)], [[':ab', ':<drawing>'], [':cd']]);
  assert.deepEqual(caret.focus, pos(2, 0));
});

test('moving by characters stops before and after a picture, and does not stop inside hidden text', async () => {
  const doc = await open(p(r('ab') + DRAWING + r('cd')) + p(DRAWING + r('x')) + p(r('ab') + r('HID', '<w:vanish/>') + r('cd')));
  const walk = (from: DocPosition, direction: 1 | -1, steps: number, unit: 'char' | 'word' = 'char', options?: { includeHidden?: boolean }): string[] => {
    const out: string[] = [];
    let at = from;
    for (let i = 0; i < steps; i++) {
      at = movePosition(doc, at, unit, direction, options);
      out.push(at.path.join('.') + (at.before ? '<' : '+') + at.offset);
    }
    return out;
  };
  assert.deepEqual(walk(pos(0, 1), 1, 7), ['0<2', '0+2', '0+3', '0+4', '1<0', '1+0', '1+1']);
  assert.deepEqual(walk(pos(1, 1), -1, 7), ['1+0', '1<0', '0+4', '0+3', '0+2', '0<2', '0+1']);
  assert.deepEqual(movePosition(doc, pos(1, 1), 'paragraphStart', -1), before(1, 0));
  assert.deepEqual(movePosition(doc, pos(0, 1), 'paragraph', 1), before(1, 0));
  // A position without a picture at its offset is given back without the flag.
  assert.deepEqual(movePosition(doc, pos(0, 3), 'char', 1), pos(0, 4));
  assert.equal('before' in movePosition(doc, pos(0, 3), 'char', 1), false);
  assert.deepEqual(walk(pos(2, 0), 1, 4), ['2+1', '2+2', '2+5', '2+6']);
  assert.deepEqual(walk(pos(2, 7), -1, 4), ['2+6', '2+5', '2+2', '2+1']);
  assert.deepEqual(walk(pos(2, 2), 1, 2, 'char', { includeHidden: true }), ['2+3', '2+4']);
  assert.deepEqual(walk(pos(2, 0), 1, 1, 'word'), ['2+7']);
});
