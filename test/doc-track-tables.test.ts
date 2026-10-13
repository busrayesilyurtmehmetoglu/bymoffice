// Tables in a document that tracks its changes: rows, columns and whole
// tables that are put in or taken out are revisions, which can be accepted
// and rejected one by one and all at once, undone, saved and read again.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { writeDocx } from '../src/formats/docx/write';
import type { DocModel, Table } from '../src/model/doc/types';
import {
  TRACK_UNSUPPORTED,
  acceptAll,
  acceptRevision,
  deleteColumn,
  deleteRange,
  deleteRow,
  insertColumn,
  insertRow,
  insertTable,
  insertText,
  rejectAll,
  rejectRevision,
  revisionsOf,
  type DocEdit,
  type DocPosition,
} from '../src/ops/doc/index';
import { checkInvariants, docPicture, p, pos, range, run, tbl, tc, texts, tr } from './doc-ops-fixtures';
import { ATTRS, ME, YOU, body, checkFile, checkReopened, open, shape } from './doc-track-fixtures';

const refused = (make: () => DocEdit): void => {
  assert.throws(make, (error: unknown) => error instanceof RangeError && error.message === TRACK_UNSUPPORTED);
};

const kinds = (doc: DocModel): string[] => revisionsOf(doc).map((revision) => `${revision.kind}:${revision.author}:${revision.rows}`);
const cell = (table: number, row: number, column: number, offset = 0): DocPosition => ({ story: 'body', path: [table, row, column, 0], offset });
const tableAt = (doc: DocModel, index: number): Table => doc.body.blocks[index] as Table;

const TABLE = tbl(tr(tc('a1') + tc('b1')) + tr(tc('a2') + tc('b2')), [1000, 1000]);
const DOC = p('before') + TABLE + p('after');

// ----------------------------------------------------------------------- rows

test('tracked: an inserted row says so in its properties, and so do the marks of its paragraphs', async () => {
  const o = await open(DOC);
  const at = run(o.history, insertRow(o.doc, cell(1, 0, 1), 'below', ME));
  assert.deepEqual(at.focus, cell(1, 1, 1));
  const table = tableAt(o.doc, 1);
  assert.equal(table.rows.length, 3);
  assert.equal(table.rows[1].props.inserted && table.rows[1].props.inserted.author, 'Me');
  const xml = await body(o);
  const mark = (id: number) => `<w:p><w:pPr><w:rPr><w:ins w:id="${id}"${ATTRS}/></w:rPr></w:pPr></w:p>`;
  assert.ok(xml.indexOf(`<w:tr><w:trPr><w:ins w:id="1"${ATTRS}/></w:trPr><w:tc>${mark(2)}</w:tc><w:tc>${mark(3)}</w:tc></w:tr>`) > 0, xml);
  await checkFile(o);
  // One change: the row with what is in it.
  assert.deepEqual(kinds(o.doc), ['insert:Me:1']);
  await checkReopened(o);
});

test('tracked: rejecting an inserted row takes it out, accepting it leaves a plain row; undo gives the bytes back', async () => {
  const o = await open(DOC);
  const before = shape(o.doc);
  const picture = docPicture(o.doc);
  run(o.history, insertRow(o.doc, cell(1, 1, 0), 'above', ME));
  run(o.history, rejectRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.equal(shape(o.doc), before, 'rejected, the table is what it was');
  assert.deepEqual(revisionsOf(o.doc), []);
  o.history.undo();
  run(o.history, acceptRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.deepEqual(revisionsOf(o.doc), []);
  assert.equal(tableAt(o.doc, 1).rows.length, 3);
  assert.equal((await body(o)).indexOf('w:ins'), -1, 'accepted, nothing says the row was inserted');
  await checkFile(o);
  while (o.history.canUndo) o.history.undo();
  assert.equal(docPicture(o.doc), picture);
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes, 'undone, the file is the one that was read');
});

test('tracked: a deleted row stays, says so, and what it holds is deleted', async () => {
  const o = await open(DOC);
  const before = shape(o.doc);
  run(o.history, deleteRow(o.doc, cell(1, 0, 0), ME));
  const table = tableAt(o.doc, 1);
  assert.equal(table.rows.length, 2, 'the row is still there');
  assert.equal(table.rows[0].props.deleted && table.rows[0].props.deleted.author, 'Me');
  assert.deepEqual(texts(o.doc), ['before', '', '', 'a2', 'b2', 'after']);
  const xml = await body(o);
  assert.ok(/<w:tr><w:trPr><w:del w:id="\d+"[^>]*\/><\/w:trPr><w:tc><w:p><w:pPr><w:rPr><w:del [^>]*\/><\/w:rPr><\/w:pPr><w:del [^>]*><w:r><w:delText>a1<\/w:delText><\/w:r><\/w:del><\/w:p><\/w:tc>/.test(xml), xml);
  await checkFile(o);
  assert.deepEqual(kinds(o.doc), ['delete:Me:1']);
  await checkReopened(o);
  // Rejected, the row is what it was; accepted, it is gone.
  run(o.history, rejectAll(o.doc));
  assert.equal(shape(o.doc), before);
  o.history.undo();
  const edit = acceptAll(o.doc);
  assert.equal(edit.skipped, 0);
  run(o.history, edit);
  assert.deepEqual(texts(o.doc), ['before', 'a2', 'b2', 'after']);
  assert.deepEqual(revisionsOf(o.doc), []);
  await checkFile(o);
});

test('tracked: deleting the row the editor inserted in this session takes it out for good; a row of the file under the same name is deleted as a change', async () => {
  const o = await open(DOC);
  const before = shape(o.doc);
  const at = run(o.history, insertRow(o.doc, cell(1, 0, 0), 'below', ME));
  run(o.history, insertText(o.doc, at.focus, 'typed', undefined, ME));
  run(o.history, deleteRow(o.doc, cell(1, 1, 0), ME));
  assert.equal(shape(o.doc), before, 'nothing is left of the row');
  assert.deepEqual(revisionsOf(o.doc), []);

  // Somebody else typed into the row: it is deleted as a tracked change, with their text in it.
  const shared = await open(DOC);
  const row = run(shared.history, insertRow(shared.doc, cell(1, 0, 0), 'below', ME));
  run(shared.history, insertText(shared.doc, row.focus, 'theirs', undefined, YOU));
  run(shared.history, deleteRow(shared.doc, cell(1, 1, 0), ME));
  assert.equal(tableAt(shared.doc, 1).rows.length, 3);
  assert.ok(tableAt(shared.doc, 1).rows[1].props.deleted, 'the row is deleted, not removed');
  assert.ok((await body(shared)).indexOf('<w:delText>theirs</w:delText>') > 0);

  const file = await open(
    p('x') + tbl(tr(tc('a1')) + `<w:tr><w:trPr><w:ins w:id="5" w:author="Me" w:date="2020-01-01T00:00:00Z"/></w:trPr>${tc('old')}</w:tr>`, [1000]) + p('y'),
  );
  run(file.history, deleteRow(file.doc, cell(1, 1, 0), ME));
  const kept = tableAt(file.doc, 1).rows[1];
  assert.ok(kept && kept.props.inserted && kept.props.deleted, 'the row of the file is inserted and deleted');
  await checkFile(file);
});

test('tracked: a row somebody deleted already is not deleted again; its cell in a column that is deleted goes with the column', async () => {
  const who = 'w:author="You" w:date="2020-01-01T00:00:00Z"';
  const gone = `<w:tr><w:trPr><w:del w:id="5" ${who}/></w:trPr>${tc('theirs')}${tc('row')}</w:tr>`;
  const o = await open(p('x') + tbl(gone + tr(tc('kept') + tc('cell')), [1000, 1000]) + p('y'));
  const before = docPicture(o.doc);
  assert.equal(deleteRow(o.doc, cell(1, 0, 0), ME).ops.length, 0, 'the deleted row is left as it is');
  assert.equal(docPicture(o.doc), before);
  run(o.history, deleteColumn(o.doc, cell(1, 1, 1), ME));
  const table = tableAt(o.doc, 1);
  // The cell of the column in the row they deleted is deleted with its column all the same: should their row come
  // back, it comes back without the column, like every other row.
  assert.equal(table.rows[0].cells[1].props.deleted && table.rows[0].cells[1].props.deleted.author, 'Me');
  assert.equal(table.rows[1].cells[1].props.deleted && table.rows[1].cells[1].props.deleted.author, 'Me');
  assert.equal(table.rows[0].props.deleted && table.rows[0].props.deleted.id, 5);
  const xml = await checkFile(o);
  assert.equal((xml.match(/<w:cellDel [^>]*w:author="Me"/g) || []).length, 2, xml);
  // The column is one change, and the row is theirs still.
  assert.deepEqual(revisionsOf(o.doc).map((revision) => `${revision.kind}:${revision.author}:${revision.columns}:${revision.rows}`), ['delete:You:0:1', 'delete:Me:1:0']);
  // All of the column deleted already: nothing more to do.
  assert.equal(deleteColumn(o.doc, cell(1, 1, 1), ME).ops.length, 0);
  o.history.undo();
  assert.equal(docPicture(o.doc), before);
});

test('tracked: the last row of a table is deleted like any other, and the table stays until that is accepted', async () => {
  const o = await open(p('x') + tbl(tr(tc('only')), [1000]) + p('y'));
  const before = shape(o.doc);
  run(o.history, deleteRow(o.doc, cell(1, 0, 0), ME));
  assert.equal(o.doc.body.blocks[1].type, 'table');
  run(o.history, acceptAll(o.doc));
  assert.deepEqual(texts(o.doc), ['x', 'y']);
  o.history.undo();
  run(o.history, rejectAll(o.doc));
  assert.equal(shape(o.doc), before);
});

test('a document that does not track inserts and deletes rows as before', async () => {
  const o = await open(DOC, { settings: '' });
  run(o.history, insertRow(o.doc, cell(1, 0, 0), 'below'));
  run(o.history, deleteRow(o.doc, cell(1, 0, 0)));
  const xml = await body(o);
  assert.equal(xml.indexOf('w:ins'), -1, xml);
  assert.equal(xml.indexOf('w:del'), -1, xml);
  assert.deepEqual(texts(o.doc), ['before', '', '', 'a2', 'b2', 'after']);
  assert.deepEqual(revisionsOf(o.doc), []);
  while (o.history.canUndo) o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);
});

test('tracked: the rows of somebody else are never taken out without a trace, and their records stay', async () => {
  const theirs = `<w:tr><w:trPr><w:ins w:id="5" w:author="You" w:date="2020-01-01T00:00:00Z"/></w:trPr>${tc('theirs')}</w:tr>`;
  const o = await open(p('x') + tbl(tr(tc('a1')) + theirs, [1000]) + p('y'));
  run(o.history, deleteRow(o.doc, cell(1, 1, 0), ME));
  run(o.history, insertRow(o.doc, cell(1, 0, 0), 'above', ME));
  const xml = await checkFile(o);
  assert.ok(xml.indexOf('<w:ins w:id="5" w:author="You" w:date="2020-01-01T00:00:00Z"/>') > 0, xml);
  // Resolving my row leaves theirs as it is.
  const mine = revisionsOf(o.doc).find((revision) => revision.kind === 'insert' && revision.author === 'Me');
  assert.ok(mine, 'the inserted row is listed');
  run(o.history, rejectRevision(o.doc, mine as NonNullable<typeof mine>));
  assert.ok(revisionsOf(o.doc).some((revision) => revision.author === 'You'), 'the row of the other author is still a change');
});

// --------------------------------------------------------------------- tables

test('tracked: an inserted table is a table of inserted rows, one change; rejected, the document is what it was', async () => {
  for (const offset of [0, 3, 6]) {
    const o = await open(p('before') + p('after'));
    const before = shape(o.doc);
    const picture = docPicture(o.doc);
    const at = run(o.history, insertTable(o.doc, pos(0, offset), 2, 2, ME));
    checkInvariants(o.doc, false);
    const index = at.focus.path[0];
    const table = tableAt(o.doc, index);
    assert.equal(table.type, 'table');
    assert.ok(table.rows.every((row) => row.props.inserted && row.props.inserted.author === 'Me'), 'every row is inserted');
    const xml = await checkFile(o, `offset ${offset}`);
    assert.equal((xml.match(/<w:trPr><w:ins /g) || []).length, 2, xml);
    assert.deepEqual(kinds(o.doc).length, 1, `offset ${offset}: ${kinds(o.doc).join(' ')}`);
    await checkReopened(o, `offset ${offset}`);
    const edit = rejectAll(o.doc);
    assert.equal(edit.skipped, 0);
    run(o.history, edit);
    assert.equal(shape(o.doc), before, `offset ${offset}: rejected`);
    o.history.undo();
    run(o.history, acceptAll(o.doc));
    assert.deepEqual(revisionsOf(o.doc), []);
    assert.equal((await body(o)).indexOf('<w:ins '), -1, 'accepted, nothing says the table was inserted');
    assert.ok(o.doc.body.blocks.some((block) => block.type === 'table'), 'accepted, the table stays');
    while (o.history.canUndo) o.history.undo();
    assert.equal(docPicture(o.doc), picture);
  }
});

test('tracked: a table put beside a table, and at the end of a cell, brings paragraphs that go with it when it is rejected', async () => {
  const o = await open(TABLE + p('after'));
  const before = shape(o.doc);
  run(o.history, insertTable(o.doc, pos(1, 0), 1, 1, ME));
  checkInvariants(o.doc, false);
  await checkFile(o);
  run(o.history, rejectAll(o.doc));
  assert.equal(shape(o.doc), before);

  const inner = await open(DOC);
  const was = shape(inner.doc);
  run(inner.history, insertTable(inner.doc, cell(1, 0, 0, 2), 1, 2, ME));
  checkInvariants(inner.doc, false);
  await checkFile(inner);
  await checkReopened(inner);
  run(inner.history, rejectAll(inner.doc));
  assert.equal(shape(inner.doc), was);
});

test('tracked: a range that takes a table deletes its rows as a change; rejected, everything is back', async () => {
  const o = await open(DOC);
  const before = shape(o.doc);
  run(o.history, deleteRange(o.doc, range(pos(0, 2), pos(2, 3)), ME));
  const table = tableAt(o.doc, 1);
  assert.equal(table.type, 'table', 'the table stays');
  assert.ok(table.rows.every((row) => row.props.deleted), 'all its rows are deleted');
  assert.deepEqual(texts(o.doc), ['be', '', '', '', '', 'er']);
  await checkFile(o);
  await checkReopened(o);
  run(o.history, rejectAll(o.doc));
  assert.equal(shape(o.doc), before);
  o.history.undo();
  const edit = acceptAll(o.doc);
  assert.equal(edit.skipped, 0);
  run(o.history, edit);
  assert.deepEqual(texts(o.doc), ['beer']);
  assert.deepEqual(revisionsOf(o.doc), []);
});

test('tracked: a range from a paragraph into a table, and between two cells, deletes what the cells hold and keeps the table', async () => {
  const o = await open(DOC);
  const before = shape(o.doc);
  run(o.history, deleteRange(o.doc, range(pos(0, 3), cell(1, 0, 1, 1)), ME));
  assert.deepEqual(texts(o.doc), ['bef', '', '', 'a2', 'b2', 'after']);
  assert.ok(tableAt(o.doc, 1).rows.every((row) => !row.props.deleted), 'no row is deleted');
  await checkFile(o);
  run(o.history, rejectAll(o.doc));
  assert.equal(shape(o.doc), before);

  const cells = await open(DOC);
  run(cells.history, deleteRange(cells.doc, range(cell(1, 0, 0, 1), cell(1, 1, 1, 1)), ME));
  assert.deepEqual(texts(cells.doc), ['before', '', '', '', '', 'after']);
  await checkFile(cells);
  await checkReopened(cells);
  run(cells.history, rejectAll(cells.doc));
  assert.equal(shape(cells.doc), before);
});

// -------------------------------------------------------------------- columns

test('tracked: an inserted column is a cell in every row that says so, and the grid as it was is kept', async () => {
  const o = await open(DOC);
  const before = shape(o.doc);
  const picture = docPicture(o.doc);
  run(o.history, insertColumn(o.doc, cell(1, 0, 0), 'right', ME));
  const table = tableAt(o.doc, 1);
  assert.deepEqual(table.grid, [1000, 1000, 1000]);
  assert.deepEqual(table.gridChange && table.gridChange.grid, [1000, 1000]);
  assert.ok(table.rows.every((row) => row.cells.length === 3 && row.cells[1].props.inserted && !row.cells[0].props.inserted && !row.cells[2].props.inserted));
  const xml = await checkFile(o);
  assert.ok(/<w:tblGrid><w:gridCol w:w="1000"\/><w:gridCol w:w="1000"\/><w:gridCol w:w="1000"\/><w:tblGridChange w:id="\d+"><w:tblGrid><w:gridCol w:w="1000"\/><w:gridCol w:w="1000"\/><\/w:tblGrid><\/w:tblGridChange><\/w:tblGrid>/.test(xml), xml);
  assert.equal((xml.match(/<w:cellIns w:id="\d+" w:author="Me" w:date="2026-10-05T10:00:00Z"\/>/g) || []).length, 2, xml);
  await checkReopened(o);
  const edit = rejectAll(o.doc);
  assert.equal(edit.skipped, 0);
  run(o.history, edit);
  assert.equal(shape(o.doc), before, 'rejected');
  o.history.undo();
  const accept = acceptAll(o.doc);
  assert.equal(accept.skipped, 0);
  run(o.history, accept);
  assert.deepEqual(revisionsOf(o.doc), []);
  const plain = await checkFile(o);
  assert.equal(plain.indexOf('cellIns'), -1, plain);
  assert.equal(plain.indexOf('tblGridChange'), -1, plain);
  assert.deepEqual(tableAt(o.doc, 1).grid, [1000, 1000, 1000]);
  while (o.history.canUndo) o.history.undo();
  assert.equal(docPicture(o.doc), picture);
});

test('tracked: a deleted column leaves its cells, which say so, with what they hold deleted', async () => {
  const o = await open(DOC);
  const before = shape(o.doc);
  run(o.history, deleteColumn(o.doc, cell(1, 1, 1), ME));
  const table = tableAt(o.doc, 1);
  assert.deepEqual(table.grid, [1000, 1000], 'the grid still has the column');
  assert.ok(table.rows.every((row) => row.cells.length === 2 && row.cells[1].props.deleted && !row.cells[0].props.deleted));
  assert.deepEqual(texts(o.doc), ['before', 'a1', '', 'a2', '', 'after']);
  const xml = await checkFile(o);
  assert.equal((xml.match(/<w:cellDel /g) || []).length, 2, xml);
  assert.ok(xml.indexOf('<w:delText>b1</w:delText>') > 0, xml);
  await checkReopened(o);
  run(o.history, rejectAll(o.doc));
  assert.equal(shape(o.doc), before, 'rejected');
  o.history.undo();
  const accept = acceptAll(o.doc);
  assert.equal(accept.skipped, 0);
  run(o.history, accept);
  assert.deepEqual(texts(o.doc), ['before', 'a1', 'a2', 'after']);
  assert.deepEqual(tableAt(o.doc, 1).grid, [1000]);
  assert.deepEqual(revisionsOf(o.doc), []);
  checkInvariants(o.doc);
  await checkFile(o);
});

test('tracked: the column the editor inserted in this session goes for good when it is deleted', async () => {
  const o = await open(DOC);
  const before = shape(o.doc);
  const at = run(o.history, insertColumn(o.doc, cell(1, 0, 1), 'left', ME));
  run(o.history, deleteColumn(o.doc, at.focus, ME));
  assert.equal(shape(o.doc), before);
  assert.deepEqual(revisionsOf(o.doc), []);
});

test('tracked: a column that cuts through merged cells is refused, and nothing changes', async () => {
  const o = await open(p('x') + tbl(tr(tc('wide', '<w:gridSpan w:val="2"/>')) + tr(tc('a2') + tc('b2')), [1000, 1000]) + p('y'));
  const before = docPicture(o.doc);
  refused(() => insertColumn(o.doc, cell(1, 1, 0), 'right', ME));
  refused(() => deleteColumn(o.doc, cell(1, 1, 0), ME));
  assert.equal(docPicture(o.doc), before);
  // A column beside the merged cell is none of its business.
  run(o.history, insertColumn(o.doc, cell(1, 1, 1), 'right', ME));
  await checkFile(o);
});

test('tracked: a table of a fixed width is as wide as its columns, pending ones included; resolving them makes it what it has to be', async () => {
  const fixed = (rows: string) => `<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>${rows}</w:tbl>`;
  const wide = (text: string) => tc(text, '<w:tcW w:w="1000" w:type="dxa"/>');
  const o = await open(p('x') + fixed(tr(wide('a1') + wide('b1')) + tr(wide('a2') + wide('b2'))) + p('y'));
  const before = shape(o.doc);
  const width = (): number | undefined => tableAt(o.doc, 1).props.width?.value;
  run(o.history, insertColumn(o.doc, cell(1, 0, 0), 'right', ME));
  assert.equal(width(), 3000, 'the new column is part of the table');
  run(o.history, rejectAll(o.doc));
  assert.equal(shape(o.doc), before, 'rejected, the table has the width it had');
  o.history.undo();
  o.history.undo();
  // A deleted column is still there until that is accepted.
  run(o.history, deleteColumn(o.doc, cell(1, 0, 1), ME));
  assert.equal(width(), 2000);
  run(o.history, acceptAll(o.doc));
  assert.equal(width(), 1000);
  assert.deepEqual(tableAt(o.doc, 1).grid, [1000]);
  await checkFile(o);
});

test('the changes of cells and of the grid a file brings are listed and resolved', async () => {
  const who = (id: number) => `w:id="${id}" w:author="Ann" w:date="2026-10-01T09:00:00Z"`;
  const grid = '<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:tblGridChange w:id="9"><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid></w:tblGridChange></w:tblGrid>';
  const row = (a: string, b: string, n: number) => `<w:tr>${tc(a)}<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/><w:cellIns ${who(n)}/></w:tcPr>${p(b)}</w:tc></w:tr>`;
  const o = await open(p('x') + `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>${grid}${row('a1', 'b1', 10)}${row('a2', 'b2', 11)}</w:tbl>` + p('y'));
  const table = tableAt(o.doc, 1);
  assert.deepEqual(table.gridChange, { id: 9, grid: [1000] });
  assert.equal(table.rows[0].cells[1].props.inserted && table.rows[0].cells[1].props.inserted.id, 10);
  const list = revisionsOf(o.doc);
  assert.ok(list.every((revision) => revision.kind !== 'other' && revision.canAccept && revision.canReject), JSON.stringify(list));
  // Untouched, the file is written as it was read.
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);
  const edit = rejectAll(o.doc);
  assert.equal(edit.skipped, 0);
  run(o.history, edit);
  assert.deepEqual(texts(o.doc), ['x', 'a1', 'a2', 'y']);
  assert.deepEqual(tableAt(o.doc, 1).grid, [1000]);
  assert.equal(tableAt(o.doc, 1).gridChange, undefined);
  assert.deepEqual(revisionsOf(o.doc), []);
  checkInvariants(o.doc);
  o.history.undo();
  const accept = acceptAll(o.doc);
  assert.equal(accept.skipped, 0);
  run(o.history, accept);
  assert.deepEqual(revisionsOf(o.doc), []);
  assert.deepEqual(tableAt(o.doc, 1).grid, [1000, 1000]);
  await checkFile(o);
});
