// The rules of tracked tables, links and pastes one by one: each test holds a
// rule that a change of the code could break without any other test seeing
// it (found by changing the code and running the tests).

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { writeDocx } from '../src/formats/docx/write';
import type { DocModel, Paragraph, Table } from '../src/model/doc/types';
import {
  REVISION_UNSUPPORTED,
  TRACK_UNSUPPORTED,
  acceptAll,
  acceptRevision,
  deleteColumn,
  deleteRange,
  deleteRow,
  fragmentOf,
  insertBlock,
  insertColumn,
  insertFragment,
  insertHyperlink,
  insertRow,
  insertTable,
  insertText,
  nextRevisionId,
  rejectAll,
  rejectRevision,
  revisionsOf,
  type DocPosition,
} from '../src/ops/doc/index';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';
import { docPicture, p, pos, r, range, run, tbl, tc, texts, tr } from './doc-ops-fixtures';
import { ME, YOU, body, checkFile, checkTables, open, shape } from './doc-track-fixtures';

const Z = 'w:author="Zed" w:date="2020-01-01T00:00:00Z"';
const cell = (table: number, row: number, column: number, offset = 0): DocPosition => ({ story: 'body', path: [table, row, column, 0], offset });
const tableAt = (doc: DocModel, index: number): Table => doc.body.blocks[index] as Table;
const refused = (make: () => unknown, message = TRACK_UNSUPPORTED): void => {
  assert.throws(make, (error: unknown) => error instanceof RangeError && error.message === message);
};
const count = (text: string, part: string): number => text.split(part).length - 1;
const TABLE = tbl(tr(tc('a1') + tc('b1')) + tr(tc('a2') + tc('b2')), [1000, 1000]);
const marked = (paragraph: unknown, flag: 'inserted' | 'deleted'): boolean => {
  const mark = (paragraph as Paragraph).props.markProps;
  return mark !== undefined && mark[flag] !== undefined;
};

// ------------------------------------------------------------ ranges and cells

test('rules: of a cell a range reaches into, the mark of the last paragraph is not deleted: the cell stays', async () => {
  const o = await open(p('before') + tbl(tr(tc(p('one') + p('two')) + tc('b1')) + tr(tc('a2') + tc('b2')), [1000, 1000]) + p('after'));
  const before = shape(o.doc);
  run(o.history, deleteRange(o.doc, range(pos(0, 2), cell(1, 0, 0, 1)), ME));
  const first = tableAt(o.doc, 1).rows[0].cells[0].blocks;
  assert.ok(marked(first[0], 'deleted'), 'the mark between the two paragraphs of the cell is deleted');
  assert.ok(!marked(first[1], 'deleted'), 'the mark that ends the cell is not');
  run(o.history, acceptAll(o.doc));
  assert.deepEqual(texts(o.doc), ['be', '', 'b1', 'a2', 'b2', 'after']);
  o.history.undo();
  run(o.history, rejectAll(o.doc));
  assert.equal(shape(o.doc), before);
  // Between two cells of one table too.
  const cells = await open(p('x') + TABLE + p('y'));
  run(cells.history, deleteRange(cells.doc, range(cell(1, 0, 0, 1), cell(1, 0, 1, 1)), ME));
  for (const one of tableAt(cells.doc, 1).rows[0].cells) assert.ok(!marked(one.blocks[0], 'deleted'), 'the cells keep their marks');
});

test('rules: a table a range covers from its very start, or to its very end, is deleted whole; one it only reaches into is not', async () => {
  const fromStart = await open(TABLE + p('after'));
  run(fromStart.history, deleteRange(fromStart.doc, range(cell(0, 0, 0, 0), pos(1, 2)), ME));
  assert.ok(tableAt(fromStart.doc, 0).rows.every((row) => row.props.deleted), 'covered from its first position: every row is deleted');
  const partFromStart = await open(TABLE + p('after'));
  run(partFromStart.history, deleteRange(partFromStart.doc, range(cell(0, 0, 0, 1), pos(1, 2)), ME));
  assert.ok(tableAt(partFromStart.doc, 0).rows.every((row) => !row.props.deleted), 'begun inside its first cell: no row is');

  const toEnd = await open(p('before') + TABLE + p('after'));
  run(toEnd.history, deleteRange(toEnd.doc, range(pos(0, 2), cell(1, 1, 1, 2)), ME));
  assert.ok(tableAt(toEnd.doc, 1).rows.every((row) => row.props.deleted), 'covered to its last position: every row is deleted');
  const partToEnd = await open(p('before') + TABLE + p('after'));
  run(partToEnd.history, deleteRange(partToEnd.doc, range(pos(0, 2), cell(1, 1, 1, 1)), ME));
  assert.ok(tableAt(partToEnd.doc, 1).rows.every((row) => !row.props.deleted), 'ended inside its last cell: no row is');
});

// ----------------------------------------------------------------- ownership

test('rules: a column the editor put in is deleted as a change, not taken out, once somebody else typed into it', async () => {
  const o = await open(p('x') + TABLE + p('y'));
  const at = run(o.history, insertColumn(o.doc, cell(1, 0, 0), 'right', ME));
  run(o.history, insertText(o.doc, at.focus, 'theirs', undefined, YOU));
  run(o.history, deleteColumn(o.doc, cell(1, 0, 1), ME));
  const table = tableAt(o.doc, 1);
  assert.ok(table.rows.every((row) => row.cells.length === 3 && row.cells[1].props.inserted && row.cells[1].props.deleted), 'the cells stay, inserted and deleted');
  assert.ok((await body(o)).indexOf('theirs') > 0, 'and the text of the other author with them');
});

test('rules: a row the editor put in is not taken out for good while a row of somebody else stands in a table inside it', async () => {
  const o = await open(p('x') + TABLE + p('y'));
  run(o.history, insertRow(o.doc, cell(1, 0, 0), 'below', ME));
  // A table of mine in the new row, and a row of theirs in that table.
  run(o.history, insertTable(o.doc, cell(1, 1, 0, 0), 1, 1, ME));
  run(o.history, insertRow(o.doc, { story: 'body', path: [1, 1, 0, 0, 0, 0, 0], offset: 0 }, 'below', YOU));
  run(o.history, deleteRow(o.doc, cell(1, 1, 1), ME));
  const table = tableAt(o.doc, 1);
  assert.equal(table.rows.length, 3, 'the row is still there');
  assert.ok(table.rows[1].props.inserted && table.rows[1].props.deleted, 'deleted as a tracked change');
  assert.ok((await body(o)).indexOf('w:author="You"') > 0, 'the row of the other author is in the file');
});

// --------------------------------------------- columns a row leaves out, records

test('rules: a tracked column where a row leaves grid columns out is refused, before the row and after it', async () => {
  const skip = (trPr: string, cells: string) => `<w:tr><w:trPr>${trPr}</w:trPr>${cells}</w:tr>`;
  const o = await open(p('x') + tbl(skip('<w:gridBefore w:val="1"/>', tc('b1') + tc('c1')) + tr(tc('a2') + tc('b2') + tc('c2')) + skip('<w:gridAfter w:val="1"/>', tc('a3') + tc('b3')), [1000, 1000, 1000]) + p('y'));
  const before = docPicture(o.doc);
  // A column at the left of the first cell of the full row lies where the first row leaves one out.
  refused(() => insertColumn(o.doc, cell(1, 1, 0), 'left', ME));
  // One at the right end lies beyond the last cell of the last row.
  refused(() => insertColumn(o.doc, cell(1, 1, 2), 'right', ME));
  // And the columns themselves cannot be deleted as a change: a row has no cell there.
  refused(() => deleteColumn(o.doc, cell(1, 1, 0), ME));
  refused(() => deleteColumn(o.doc, cell(1, 1, 2), ME));
  assert.equal(docPicture(o.doc), before);
  // The column every row has a cell in is like any other.
  run(o.history, deleteColumn(o.doc, cell(1, 1, 1), ME));
  checkTables(o.doc);
  await checkFile(o);
});

test('rules: the record of the grid is made once, and the editor takes back only one of their own', async () => {
  const o = await open(p('x') + TABLE + p('y'));
  run(o.history, insertColumn(o.doc, cell(1, 0, 0), 'right', ME));
  const record = tableAt(o.doc, 1).gridChange;
  assert.deepEqual(record && record.grid, [1000, 1000]);
  run(o.history, insertColumn(o.doc, cell(1, 0, 2), 'right', ME));
  assert.deepEqual(tableAt(o.doc, 1).gridChange, record, 'a second column leaves the record what it was: the grid before the first');
  // Both taken back: the grid is what the record says, and the record of this session goes.
  run(o.history, deleteColumn(o.doc, cell(1, 0, 3), ME));
  assert.ok(tableAt(o.doc, 1).gridChange, 'one column is still pending: the record stays');
  run(o.history, deleteColumn(o.doc, cell(1, 0, 1), ME));
  assert.equal(tableAt(o.doc, 1).gridChange, undefined);
  assert.deepEqual(revisionsOf(o.doc), []);

  // A record somebody else left, whose grid the table has again, is theirs: it stays.
  const theirs = '<w:tblGridChange w:id="90"><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid></w:tblGridChange>';
  const file = await open(p('x') + `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/>${theirs}</w:tblGrid>${tr(tc('a1') + tc('b1'))}</w:tbl>` + p('y'));
  const at = run(file.history, insertColumn(file.doc, cell(1, 0, 0), 'right', ME));
  run(file.history, deleteColumn(file.doc, at.focus, ME));
  assert.deepEqual(tableAt(file.doc, 1).gridChange, { id: 90, grid: [1000, 1000] });
});

// ------------------------------------------------ content controls in a change

const CONTROL = (inner: string) => `<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;

test('rules: a link at a caret, and a paste, inside a content control that sits in the change of somebody else is refused', async () => {
  const o = await open(p(`<w:ins w:id="50" ${Z}>${CONTROL(r('abcd'))}</w:ins>`) + p('plain'));
  const before = docPicture(o.doc);
  const relationships = o.doc.body.relationships;
  refused(() => insertHyperlink(o.doc, range(pos(0, 2), pos(0, 2)), 'https://example.org/', ME));
  const source = await open(p(r('one ') + r('two', '<w:b/>')));
  const fragment = fragmentOf(source.doc, range(pos(0, 1), pos(0, 6)));
  refused(() => insertFragment(o.doc, pos(0, 2), fragment, ME));
  assert.equal(docPicture(o.doc), before);
  assert.ok(o.doc.body.relationships === relationships);
  // A control that is in no change is no obstacle.
  const free = await open(p(CONTROL(r('abcd'))));
  run(free.history, insertFragment(free.doc, pos(0, 2), fragment, ME));
  assert.equal(count(await checkFile(free), '<w:sdt>'), 1);
});

test('rules: what is pasted where the caret stands in the insertion of somebody else is no insertion inside theirs', async () => {
  const o = await open(p(`<w:ins w:id="50" ${Z}>${r('abcd')}</w:ins>`));
  const source = await open(p(r('one ') + r('two', '<w:b/>')));
  run(o.history, insertFragment(o.doc, pos(0, 2), fragmentOf(source.doc, range(pos(0, 1), pos(0, 6))), ME));
  const xml = await checkFile(o);
  assert.ok(!/<w:ins [^>]*>(?:(?!<\/w:ins>).)*<w:ins /.test(xml), xml);
  assert.equal(count(xml, 'w:author="Zed"'), 2, 'theirs is cut in two');
  assert.equal(count(xml, 'w:author="Me"'), 1);
});

test('rules: a block put in whole does not put what is deleted in it into an insertion', async () => {
  const o = await open(p('abcd'));
  const paragraph: Paragraph = {
    type: 'paragraph',
    props: {},
    children: [
      { type: 'run', props: {}, content: [{ type: 'text', text: 'new' }] },
      { type: 'run', props: {}, content: [{ type: 'deletedText', text: 'old' }], wrap: { kind: 'del', id: 900, author: 'Zed', date: '2020-01-01T00:00:00Z' } },
    ],
  };
  run(o.history, insertBlock(o.doc, pos(0, 0), paragraph, ME));
  const xml = await checkFile(o);
  assert.ok(/<w:del w:id="900"[^>]*><w:r><w:delText>old<\/w:delText><\/w:r><\/w:del>/.test(xml), xml);
  assert.equal(count(xml, '<w:ins '), 2, 'the text and the mark, nothing else');
});

// ---------------------------------------------------------------------- paste

test('rules: a paste that begins with a table cuts the paragraph with an inserted mark, and a paragraph it needs beside a table is inserted too', async () => {
  const source = await open(TABLE + p('tail'));
  const fragment = fragmentOf(source.doc, range(cell(0, 0, 0, 0), pos(1, 2)));
  // In the middle of a paragraph: the first half ends with a mark that is new.
  const mid = await open(p('abcd') + p('end'));
  const start = shape(mid.doc);
  run(mid.history, insertFragment(mid.doc, pos(0, 2), fragment, ME));
  assert.ok(marked(mid.doc.body.blocks[0], 'inserted'), 'the mark of the half before the table is inserted');
  run(mid.history, rejectAll(mid.doc));
  assert.equal(shape(mid.doc), start);
  // At the start of a paragraph after a table: a paragraph has to stand between the two tables, and it is part of the paste.
  const beside = await open(TABLE + p('abcd'));
  const was = shape(beside.doc);
  run(beside.history, insertFragment(beside.doc, pos(1, 0), fragment, ME));
  const filler = beside.doc.body.blocks[1];
  assert.ok(filler.type === 'paragraph' && filler.children.length === 0 && marked(filler, 'inserted'), 'the paragraph between the tables is inserted');
  assert.equal(beside.doc.body.blocks[2].type, 'table');
  const all = rejectAll(beside.doc);
  assert.equal(all.skipped, 0);
  run(beside.history, all);
  assert.equal(shape(beside.doc), was);
});

test('rules: a cut that is pasted back where changes are tracked changes no formatting of the paragraph it lands in', async () => {
  const o = await open(p('abcd', '<w:jc w:val="center"/>') + p('efgh', '<w:jc w:val="right"/>') + p('end'));
  const start = shape(o.doc);
  const span = range(pos(0, 2), pos(1, 2));
  const fragment = fragmentOf(o.doc, span);
  const at = run(o.history, deleteRange(o.doc, span, ME));
  run(o.history, insertFragment(o.doc, at.focus, fragment, { ...ME, moved: true }));
  const aligned = o.doc.body.blocks.filter((block) => block.type === 'paragraph').map((block) => (block as Paragraph).props.align);
  // The paragraph the paste cut and the one it continues are centred like the paragraph they were, the one that was
  // the second is still right-aligned: nothing was given "back".
  assert.deepEqual(aligned, ['center', 'center', 'right', undefined]);
  assert.equal(count(await checkFile(o), 'pPrChange'), 0);
  run(o.history, rejectAll(o.doc));
  assert.equal(shape(o.doc), start);
});

// ------------------------------------------------- what stays of what goes

const BOOKMARK = '<w:bookmarkStart w:id="700" w:name="here"/><w:bookmarkEnd w:id="700"/>';

test('rules: the bookmark in a table that goes whole is kept in a paragraph of its own, and the paragraph before does not join past it', async () => {
  const row = `<w:tr><w:trPr><w:ins w:id="61" ${Z}/></w:trPr>${tc(p(BOOKMARK + r('cell')))}</w:tr>`;
  const o = await open(p('ab', `<w:rPr><w:ins w:id="60" ${Z}/></w:rPr>`) + tbl(row, [1000]) + p('cd'));
  const all = rejectAll(o.doc);
  run(o.history, all);
  const xml = await checkFile(o);
  assert.equal(count(xml, '<w:bookmarkStart '), 1, xml);
  assert.equal(count(xml, '<w:bookmarkEnd '), 1, xml);
  assert.deepEqual(texts(o.doc), ['ab', '', 'cd']);
  assert.equal(count(xml, '<w:tbl>'), 0);
  // The mark does not wait for a table that goes with it: nothing is left to resolve.
  assert.deepEqual([all.resolved, all.skipped], [1, 0]);
  assert.deepEqual(revisionsOf(o.doc), []);
});

test('rules: a bookmark between the paragraphs of a cell stays when its table goes and the paragraphs around the table are joined', async () => {
  const row = `<w:tr><w:trPr><w:ins w:id="61" ${Z}/></w:trPr><w:tc>${p('one')}${BOOKMARK}${p('two')}</w:tc></w:tr>`;
  const o = await open(p('ab', `<w:rPr><w:ins w:id="60" ${Z}/></w:rPr>`) + tbl(row, [1000]) + p('cd'));
  run(o.history, rejectAll(o.doc));
  const xml = await checkFile(o);
  assert.deepEqual(texts(o.doc), ['abcd']);
  assert.equal(count(xml, '<w:bookmarkStart '), 1, xml);
  assert.equal(count(xml, '<w:bookmarkEnd '), 1, xml);
});

test('rules: the bookmark in a cell that goes with its column moves to the cell beside it', async () => {
  const gone = (id: number, inner: string) => `<w:tc><w:tcPr><w:cellDel w:id="${id}" ${Z}/></w:tcPr>${inner}</w:tc>`;
  const o = await open(p('x') + tbl(tr(tc('a1') + gone(61, p(BOOKMARK + r('b1')))) + tr(tc('a2') + gone(62, p('b2'))), [1000, 1000]) + p('y'));
  const all = acceptAll(o.doc);
  assert.equal(all.skipped, 0);
  run(o.history, all);
  const xml = await checkFile(o);
  assert.equal(count(xml, '<w:bookmarkStart '), 1, xml);
  assert.deepEqual(texts(o.doc), ['x', 'a1', 'a2', 'y']);
  checkTables(o.doc);
});

test('rules: a column whose going would take half a field is not resolved', async () => {
  const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> QUOTE x </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>res</w:t></w:r>';
  const end = '<w:r><w:t>ult</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const gone = (id: number, inner: string) => `<w:tc><w:tcPr><w:cellDel w:id="${id}" ${Z}/></w:tcPr>${inner}</w:tc>`;
  // The field begins in a cell of the deleted column and ends in the cell after it.
  const o = await open(p('x') + tbl(tr(gone(61, p(begin)) + tc(p(end))) + tr(gone(62, p('a2')) + tc('b2')), [1000, 1000]) + p('y'));
  const before = shape(o.doc);
  const all = acceptAll(o.doc);
  assert.deepEqual([all.resolved, all.skipped, all.ops.length], [0, 1, 0]);
  refused(() => acceptRevision(o.doc, revisionsOf(o.doc)[0]), REVISION_UNSUPPORTED);
  assert.equal(shape(o.doc), before);
  // Rejecting takes nothing out: that is done.
  const rejected = rejectAll(o.doc);
  assert.deepEqual([rejected.resolved, rejected.skipped], [1, 0]);
});

// ------------------------------------------------- formatting of a cell (kept)

test('rules: rejecting a change of the formatting of a cell leaves the mark of its column, and the columns it covers', async () => {
  const record = `<w:tcPrChange w:id="70" w:author="Ann" w:date="2020-01-01T00:00:00Z"><w:tcPr><w:tcW w:w="900" w:type="dxa"/></w:tcPr></w:tcPrChange>`;
  for (const kind of ['cellIns', 'cellDel'] as const) {
    const cellOf = (id: number, text: string, extra = '') => `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:${kind} w:id="${id}" ${Z}/>${extra}</w:tcPr>${p(text)}</w:tc>`;
    const o = await open(p('x') + tbl(tr(tc('a1') + cellOf(61, 'b1', record)) + tr(tc('a2') + cellOf(62, 'b2')), [1000, 2000]) + p('y'));
    const format = revisionsOf(o.doc).find((revision) => revision.kind === 'tableFormat');
    assert.ok(format && format.canReject, kind);
    run(o.history, rejectRevision(o.doc, format));
    const props = tableAt(o.doc, 1).rows[0].cells[1].props;
    assert.equal(props.width && props.width.value, 900, kind);
    assert.equal((kind === 'cellIns' ? props.inserted : props.deleted)?.id, 61, `${kind}: the mark of the column is still there`);
    assert.deepEqual(revisionsOf(o.doc).map((revision) => `${revision.kind}:${revision.columns}`), [`${kind === 'cellIns' ? 'insert' : 'delete'}:1`], kind);
  }
  // A cell over two columns whose record names one: the cell covers what it covers now.
  const wide = `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:gridSpan w:val="2"/><w:tcPrChange w:id="70" ${Z}><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr></w:tcPrChange></w:tcPr>${p('wide')}</w:tc>`;
  const o = await open(p('x') + tbl(tr(wide) + tr(tc('a2') + tc('b2')), [1000, 1000]) + p('y'));
  run(o.history, rejectAll(o.doc));
  assert.equal(tableAt(o.doc, 1).rows[0].cells[0].props.gridSpan, 2);
  checkTables(o.doc);
});

// ------------------------------------------------------- comments, ids, reader

test('rules: the changes of cells and of the grid inside a comment are counted with the changes that are only kept', async () => {
  const grid = '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:tblGridChange w:id="80"><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid></w:tblGridChange></w:tblGrid>';
  const row = `<w:tr>${tc('a')}<w:tc><w:tcPr><w:cellIns w:id="81" ${Z}/><w:cellDel w:id="82" ${Z}/></w:tcPr>${p('b')}</w:tc></w:tr>`;
  const comment = `<w:comment w:id="1" w:author="Ann"><w:tbl><w:tblPr/>${grid}${row}</w:tbl><w:p/></w:comment>`;
  const o = await open(p('<w:commentRangeStart w:id="1"/>' + r('text') + '<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r>'), { comments: comment });
  const list = revisionsOf(o.doc);
  assert.deepEqual(list.map((revision) => [revision.kind, revision.count]), [['other', 3]]);
});

test('rules: a new revision gets an id beyond those of cells and of the record of a grid, also where the part has no source to search', async () => {
  const grid = '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:tblGridChange w:id="950"><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid></w:tblGridChange></w:tblGrid>';
  const make = async (cellId: number, withGrid: boolean) => {
    const row = `<w:tr>${tc('a')}<w:tc><w:tcPr><w:cellIns w:id="${cellId}" ${Z}/></w:tcPr>${p('b')}</w:tc></w:tr>`;
    const o = await open(p('x') + `<w:tbl><w:tblPr/>${withGrid ? grid : '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>'}${row}</w:tbl>` + p('y'));
    // As a document read from a format without XML: the model is all there is.
    o.doc.body.source = '';
    return o;
  };
  assert.equal(nextRevisionId((await make(700, false)).doc), 701, 'the inserted cell');
  assert.equal(nextRevisionId((await make(700, true)).doc), 951, 'the record of the grid');
  const deleted = await open(p('x') + tbl(`<w:tr>${tc('a')}<w:tc><w:tcPr><w:cellDel w:id="800" ${Z}/></w:tcPr>${p('b')}</w:tc></w:tr>`, [1000, 1000]) + p('y'));
  deleted.doc.body.source = '';
  assert.equal(nextRevisionId(deleted.doc), 801, 'the deleted cell');
});

test('rules: the reader takes a record of the grid only when it is the last thing in the grid and carries an id and nothing else', async () => {
  const was = (attrs: string) => `<w:tblGridChange ${attrs}><w:tblGrid><w:gridCol w:w="500"/></w:tblGrid></w:tblGridChange>`;
  const doc = async (inner: string) => (await open(p('x') + `<w:tbl><w:tblPr/><w:tblGrid>${inner}</w:tblGrid>${tr(tc('a') + tc('b'))}</w:tbl>` + p('y'))).doc;
  const plain = tableAt(await doc('<w:gridCol w:w="1000"/><w:gridCol w:w="1000"/>' + was('w:id="5"')), 1);
  assert.deepEqual([plain.gridChange, plain.extras], [{ id: 5, grid: [500] }, undefined]);
  // A column after the record: the grid is not what the schema says, and is kept as it was read.
  const followed = await doc('<w:gridCol w:w="1000"/>' + was('w:id="5"') + '<w:gridCol w:w="1000"/>');
  const follows = tableAt(followed, 1);
  assert.equal(follows.gridChange, undefined);
  assert.deepEqual(follows.grid, [1000, 1000]);
  assert.equal(follows.extras && follows.extras.length, 1);
  // A record that says more than an id.
  const named = await doc('<w:gridCol w:w="1000"/><w:gridCol w:w="1000"/>' + was('w:id="5" w:author="Zed"'));
  const signed = tableAt(named, 1);
  assert.equal(signed.gridChange, undefined);
  assert.equal(signed.extras && signed.extras.length, 1);
  // Both are listed as changes that are only kept.
  for (const model of [followed, named]) assert.deepEqual(revisionsOf(model).map((revision) => revision.kind), ['other']);
});

test('rules: of two marks of the same kind in the properties of a cell the first counts, and the second is kept as it is', async () => {
  const o = await open(p('x') + tbl(`<w:tr>${tc('a')}<w:tc><w:tcPr><w:cellIns w:id="5" ${Z}/><w:cellIns w:id="6" w:author="Ann"/></w:tcPr>${p('b')}</w:tc></w:tr>`, [1000, 1000]) + p('y'));
  const props = tableAt(o.doc, 1).rows[0].cells[1].props;
  assert.deepEqual(props.inserted, { id: 5, author: 'Zed', date: '2020-01-01T00:00:00Z' });
  assert.deepEqual((props.extras || []).map((extra) => extra.el.attrs.get('w:id')), ['6']);
});

test('rules: a grid that is kept as read keeps its record with it: the record is one change, which is only kept, and is written as it was read', async () => {
  // A column with an attribute of another vocabulary: the grid is kept, and its record in it.
  const grid = '<w:tblGrid><w:gridCol w:w="1000" x:odd="1" xmlns:x="urn:x"/><w:gridCol w:w="1000"/><w:tblGridChange w:id="5"><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid></w:tblGridChange></w:tblGrid>';
  const o = await open(p('x') + `<w:tbl><w:tblPr/>${grid}${tr(tc('a') + tc('b'))}</w:tbl>` + p('y'), { settings: '' });
  assert.equal(tableAt(o.doc, 1).gridChange, undefined);
  assert.deepEqual(revisionsOf(o.doc).map((revision) => [revision.kind, revision.count]), [['other', 1]]);
  // The table is written from the model (text typed into a cell): the grid comes out as it went in, once.
  run(o.history, insertText(o.doc, cell(1, 0, 0, 1), 'Z'));
  await writeDocx(o.pkg, o.doc);
  const xml = await body(o);
  assert.ok(xml.indexOf(grid) > 0, xml);
  assert.equal(count(xml, '<w:tblGridChange '), 1, xml);
  const all = acceptAll(o.doc);
  assert.deepEqual([all.resolved, all.skipped, all.ops.length], [0, 1, 0]);
  // A tracked column in such a table makes no second record: the grid is written from what is kept.
  run(o.history, insertColumn(o.doc, cell(1, 0, 0), 'right', { ...ME, trackChanges: true }));
  assert.ok(tableAt(o.doc, 1).rows[0].cells[1].props.inserted, 'the column is tracked');
  assert.equal(tableAt(o.doc, 1).gridChange, undefined);
  const after = await checkFile(o);
  assert.equal(count(after, '<w:tblGridChange '), 1, after);
  assert.equal(count(after, '<w:gridCol '), 4, after);
});

// ------------------------------------------------- columns, waiting, wording

test('rules: an inserted column that is rejected leaves the grid the table has without it, not the widths of the record', async () => {
  const cellIns = (id: number, text: string) => `<w:tc><w:tcPr><w:tcW w:w="700" w:type="dxa"/><w:cellIns w:id="${id}" ${Z}/></w:tcPr>${p(text)}</w:tc>`;
  // The record says the two columns were 1500 wide each; they are 1000 and 2000 now.
  const grid = '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="700"/><w:gridCol w:w="2000"/><w:tblGridChange w:id="90"><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid></w:tblGridChange></w:tblGrid>';
  const o = await open(p('x') + `<w:tbl><w:tblPr><w:tblW w:w="3700" w:type="dxa"/></w:tblPr>${grid}<w:tr>${tc('a')}${cellIns(61, 'n')}${tc('b')}</w:tr></w:tbl>` + p('y'));
  run(o.history, rejectAll(o.doc));
  const table = tableAt(o.doc, 1);
  assert.deepEqual([table.grid, table.gridChange, table.props.width && table.props.width.value], [[1000, 2000], undefined, 3000]);
  checkTables(o.doc);
});

test('rules: a change that is nothing but a mark that has to wait is not said to be resolved', async () => {
  const row = `<w:tr><w:trPr><w:ins w:id="61" w:author="Ann" w:date="2020-01-01T00:00:00Z"/></w:trPr>${tc('cell')}</w:tr>`;
  const o = await open(p('ab', `<w:rPr><w:ins w:id="60" ${Z}/></w:rPr>`) + tbl(row, [1000]) + p('cd'));
  const before = shape(o.doc);
  // Zed's changes alone: the mark, which waits for Ann's table.
  const zed = rejectAll(o.doc, (revision) => revision.author === 'Zed');
  assert.deepEqual([zed.resolved, zed.skipped, zed.ops.length], [0, 1, 0]);
  assert.equal(shape(o.doc), before);
  const all = rejectAll(o.doc);
  assert.deepEqual([all.resolved, all.skipped], [2, 0]);
  run(o.history, all);
  assert.deepEqual(texts(o.doc), ['abcd']);
});

test('rules: a column is said to be a column, in English and in Turkish', () => {
  assert.equal(STRINGS_EN.revisionDescription('insertColumn', 'Ann', '2026-10-01 09:00', ''), 'Column inserted by Ann, 2026-10-01 09:00');
  assert.equal(STRINGS_EN.revisionDescription('deleteColumn', undefined, undefined, 'b1'), 'Column deleted: b1');
  assert.equal(STRINGS_TR.revisionDescription('insertColumn', 'Ann', undefined, ''), 'Sütun eklendi (Ann)');
  assert.equal(STRINGS_TR.revisionDescription('deleteColumn', undefined, undefined, ''), 'Sütun silindi');
});
