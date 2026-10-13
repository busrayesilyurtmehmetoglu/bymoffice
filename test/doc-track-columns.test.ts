// A column that was put in or taken out is one tracked change: one entry of
// the list, resolved with all its cells and with the record of the grid. No
// order of resolving changes one by one leaves a table whose rows do not fit
// its grid, and every order leaves what resolving them all at once leaves.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import type { DocModel, Table } from '../src/model/doc/types';
import {
  TRACK_UNSUPPORTED,
  acceptAll,
  acceptRevision,
  deleteColumn,
  deleteRow,
  insertColumn,
  insertRow,
  insertText,
  listRevisions,
  rejectAll,
  rejectRevision,
  revisionsOf,
  revisionsTouching,
  type DocEdit,
  type DocPosition,
} from '../src/ops/doc/index';
import { checkInvariants, p, run, tc, texts, tr } from './doc-ops-fixtures';
import { ME, YOU, body, checkFile, checkReopened, checkTables, open, shape, type Opened } from './doc-track-fixtures';

const W = (n: number) => `<w:tcW w:w="${n}" w:type="dxa"/>`;
const table = (rows: string, grid: number[], width = grid.reduce((sum, w) => sum + w, 0), inGrid = '') =>
  `<w:tbl><w:tblPr><w:tblW w:w="${width}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}${inGrid}</w:tblGrid>${rows}</w:tbl>`;
const FIXED = p('alpha beta') + table(tr(tc('one', W(1000)) + tc('two', W(2000)) + tc('tri', W(1500))) + tr(tc('four', W(1000)) + tc('five', W(2000)) + tc('six', W(1500))), [1000, 2000, 1500]) + p('omega end');
const cell = (row: number, column: number, offset = 0, at = 1): DocPosition => ({ story: 'body', path: [at, row, column, 0], offset });
const tableOf = (doc: DocModel, index = 1): Table => doc.body.blocks[index] as Table;
const Z = 'w:author="Zed" w:date="2020-01-01T00:00:00Z"';

/** Every order of resolving the changes of a document one by one (the k-th of those left, each time) leaves what "all" leaves. */
async function everyOrder(make: () => Promise<Opened>, verdict: 'accept' | 'reject', what: string): Promise<number> {
  const all = await make();
  const count = revisionsOf(all.doc).length;
  const edit = verdict === 'accept' ? acceptAll(all.doc) : rejectAll(all.doc);
  assert.equal(edit.skipped, 0, `${what}: "all" passes changes over`);
  run(all.history, edit);
  checkTables(all.doc, `${what} all`);
  const want = shape(all.doc);
  const wantXml = await body(all);
  // The first choices are tried in every way; after them the first change left is taken each time.
  const free = Math.min(count, 4);
  const order = new Array<number>(count).fill(0);
  let orders = 0;
  const walk = async (depth: number): Promise<void> => {
    if (depth < free) {
      for (let k = 0; k < count - depth; k++) {
        order[depth] = k;
        await walk(depth + 1);
      }
      return;
    }
    orders++;
    const o = await make();
    for (let step = 0, guard = 0; guard < 64; guard++) {
      const list = revisionsOf(o.doc);
      if (!list.length) break;
      const pick = list[Math.min(order[Math.min(step, count - 1)], list.length - 1)];
      let one: DocEdit | undefined;
      try {
        one = verdict === 'accept' ? acceptRevision(o.doc, pick) : rejectRevision(o.doc, pick);
      } catch (error) {
        assert.ok(error instanceof RangeError, String(error));
      }
      // A change that has to wait for another: the one after it, and this one comes back.
      if (!one) {
        const other = list.find((revision) => {
          try {
            one = verdict === 'accept' ? acceptRevision(o.doc, revision) : rejectRevision(o.doc, revision);
            return true;
          } catch {
            return false;
          }
        });
        assert.ok(other && one, `${what} order ${order.join('')}: no change can be resolved`);
      } else step++;
      run(o.history, one as DocEdit);
      checkTables(o.doc, `${what} order ${order.join('')}`);
      checkInvariants(o.doc, false);
    }
    assert.deepEqual(revisionsOf(o.doc), [], `${what} order ${order.join('')}: changes are left`);
    assert.ok(shape(o.doc) === want, `${what} order ${order.join('')}: the model is not the one "all" leaves`);
    assert.equal(await body(o), wantXml, `${what} order ${order.join('')}: the file is not the one "all" leaves`);
  };
  await walk(0);
  return orders;
}

const summary = (doc: DocModel): string[] => revisionsOf(doc).map((revision) => `${revision.kind}:${revision.author}:${revision.columns}:${revision.rows}`);

// ----------------------------------------------------------------------- list

test('column: an inserted column is one change, with the record of the grid in it', async () => {
  const o = await open(FIXED);
  const before = shape(o.doc);
  run(o.history, insertColumn(o.doc, cell(1, 0), 'left', ME));
  assert.deepEqual(summary(o.doc), ['insert:Me:1:0']);
  const one = revisionsOf(o.doc)[0];
  assert.deepEqual([one.canAccept, one.canReject], [true, true]);
  // It is where its cells are, one stretch for each row; a caret in one of them is in the column, one in the cell
  // between them (the next cell of the first row) is not.
  assert.deepEqual([one.start.path, one.end.path], [[1, 0, 0, 0], [1, 0, 0, 0]]);
  assert.deepEqual(one.cells && one.cells.map((stretch) => stretch.start.path), [[1, 0, 0, 0], [1, 1, 0, 0]]);
  const caret = (path: number[]) => ({ anchor: { story: 'body', path, offset: 0 }, focus: { story: 'body', path, offset: 0 } });
  run(o.history, insertText(o.doc, cell(0, 0), 'x', undefined, YOU));
  const list = revisionsOf(o.doc);
  const column = list.findIndex((revision) => revision.columns === 1);
  assert.deepEqual(revisionsTouching(list, { anchor: { story: 'body', path: [1, 1, 0, 0], offset: 0 }, focus: { story: 'body', path: [1, 1, 0, 0], offset: 0 } }), [column], 'the caret in its cell of the second row');
  assert.deepEqual(revisionsTouching(list, caret([1, 0, 1, 0])), [], 'the caret in the cell beside it');
  assert.deepEqual(revisionsTouching(list, { anchor: { story: 'body', path: [0], offset: 0 }, focus: { story: 'body', path: [2], offset: 0 } }), [list.findIndex((revision) => revision.author === 'You'), column].sort(), 'a selection over the table');
  o.history.undo();
  run(o.history, rejectRevision(o.doc, one));
  assert.equal(shape(o.doc), before, 'rejected: the table, its grid and its width are what they were');
  assert.deepEqual(revisionsOf(o.doc), []);
  o.history.undo();
  run(o.history, acceptRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.deepEqual(revisionsOf(o.doc), []);
  const t = tableOf(o.doc);
  assert.deepEqual([t.grid, t.gridChange, t.props.width && t.props.width.value], [[1000, 1000, 2000, 1500], undefined, 5500]);
  checkTables(o.doc);
  await checkFile(o);
});

test('column: a deleted column is one change; accepted, that column leaves the grid and the table is that much narrower', async () => {
  const o = await open(FIXED);
  const before = shape(o.doc);
  run(o.history, deleteColumn(o.doc, cell(0, 0), ME));
  assert.deepEqual(summary(o.doc), ['delete:Me:1:0']);
  run(o.history, acceptRevision(o.doc, revisionsOf(o.doc)[0]));
  const t = tableOf(o.doc);
  assert.deepEqual([t.grid, t.props.width && t.props.width.value], [[2000, 1500], 3500]);
  assert.deepEqual(texts(o.doc), ['alpha beta', 'two', 'tri', 'five', 'six', 'omega end']);
  assert.ok(t.rows.every((row) => row.props.gridAfter === undefined && row.cells.length === 2), 'no row is shifted or leaves a column out');
  checkTables(o.doc);
  o.history.undo();
  run(o.history, rejectRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.equal(shape(o.doc), before);
});

test('column: text typed into a new column and a row put in below belong to their own changes, and the column stays one', async () => {
  const o = await open(FIXED);
  const before = shape(o.doc);
  const at = run(o.history, insertColumn(o.doc, cell(0, 1), 'right', ME));
  run(o.history, insertText(o.doc, at.focus, 'theirs', undefined, YOU));
  run(o.history, insertRow(o.doc, cell(1, 0), 'below', YOU));
  // The cell of the new row in the pending column is part of that column: it carries its mark.
  const added = tableOf(o.doc).rows[2].cells[2];
  assert.equal(added.props.inserted && added.props.inserted.author, 'Me', 'the new row has a cell of the inserted column');
  assert.deepEqual(summary(o.doc), ['insert:Me:1:0', 'insert:You:0:0', 'insert:You:0:1']);
  await checkFile(o);
  await checkReopened(o);
  // The column rejected alone: every row loses that cell, the new row too.
  const column = revisionsOf(o.doc)[0];
  run(o.history, rejectRevision(o.doc, column));
  checkTables(o.doc);
  assert.ok(tableOf(o.doc).rows.every((row) => row.cells.length === 3));
  run(o.history, rejectAll(o.doc));
  assert.equal(shape(o.doc), before);
});

// -------------------------------------------------------------- every order

test('column: every order of rejecting, and of accepting, one by one leaves what "all" leaves', async () => {
  const scenes: [string, (o: Opened) => void][] = [
    ['insert a column', (o) => void run(o.history, insertColumn(o.doc, cell(0, 1), 'left', ME))],
    ['delete a column', (o) => void run(o.history, deleteColumn(o.doc, cell(0, 1), ME))],
    [
      'insert a column, delete another',
      (o) => {
        run(o.history, insertColumn(o.doc, cell(0, 0), 'right', ME));
        run(o.history, deleteColumn(o.doc, cell(0, 2), ME));
      },
    ],
    [
      'two columns, a row and text of another author',
      (o) => {
        run(o.history, insertColumn(o.doc, cell(0, 0), 'left', ME));
        run(o.history, insertColumn(o.doc, cell(1, 3), 'right', ME));
        run(o.history, insertRow(o.doc, cell(0, 1), 'below', YOU));
        run(o.history, insertText(o.doc, cell(0, 0), 'x', undefined, YOU));
        run(o.history, deleteRow(o.doc, cell(2, 0), ME));
      },
    ],
  ];
  let orders = 0;
  for (const [name, setup] of scenes) {
    for (const verdict of ['reject', 'accept'] as const) {
      orders += await everyOrder(
        async () => {
          const o = await open(FIXED);
          setup(o);
          return o;
        },
        verdict,
        `${name}, ${verdict}`,
      );
    }
  }
  assert.ok(orders > 40, `${orders} orders were tried`);
});

test('column: one inserted at the left and rejected leaves no row shifted; a column deleted afterwards is that column', async () => {
  const o = await open(FIXED);
  const before = shape(o.doc);
  run(o.history, insertColumn(o.doc, cell(0, 0), 'left', ME));
  run(o.history, rejectRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.equal(shape(o.doc), before);
  run(o.history, deleteColumn(o.doc, cell(0, 0), ME));
  run(o.history, acceptAll(o.doc));
  assert.deepEqual(texts(o.doc), ['alpha beta', 'two', 'tri', 'five', 'six', 'omega end']);
  assert.deepEqual(tableOf(o.doc).grid, [2000, 1500]);
  checkTables(o.doc);
  checkInvariants(o.doc);
});

// ----------------------------------------------------------------- from a file

const ins = (id: number, text: string, width: number, who = Z) => `<w:tc><w:tcPr>${W(width)}<w:cellIns w:id="${id}" ${who}/></w:tcPr>${p(text)}</w:tc>`;
const del = (id: number, text: string, width: number, who = Z) => `<w:tc><w:tcPr>${W(width)}<w:cellDel w:id="${id}" ${who}/></w:tcPr>${p(text)}</w:tc>`;
const WAS = (grid: number[]) => `<w:tblGridChange w:id="90"><w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid></w:tblGridChange>`;

test('column: the cells a file brings for one column are one change, and so is its record of the grid', async () => {
  const rows = tr(tc('a1', W(1000)) + ins(11, 'n1', 2000) + del(21, 'd1', 1500)) + tr(tc('a2', W(1000)) + ins(12, 'n2', 2000) + del(22, 'd2', 1500));
  const make = () => open(p('x') + table(rows, [1000, 2000, 1500], 4500, WAS([1000, 1500])) + p('y'));
  const o = await make();
  assert.deepEqual(summary(o.doc), ['insert:Zed:1:0', 'delete:Zed:1:0']);
  const rejected = rejectAll(o.doc);
  assert.deepEqual([rejected.resolved, rejected.skipped], [2, 0]);
  run(o.history, rejected);
  const t = tableOf(o.doc);
  assert.deepEqual([t.grid, t.gridChange, t.props.width && t.props.width.value], [[1000, 1500], undefined, 2500]);
  assert.deepEqual(texts(o.doc), ['x', 'a1', 'd1', 'a2', 'd2', 'y']);
  checkTables(o.doc);
  assert.equal((await body(o)).indexOf('tblGridChange'), -1);
  for (const verdict of ['reject', 'accept'] as const) assert.equal(await everyOrder(make, verdict, `file, ${verdict}`), 2);
});

test('column: cells that do not make a whole column are kept, listed as "other" and never resolved', async () => {
  const cases: [string, string][] = [
    ['in one row only', tr(tc('a1', W(1000)) + ins(11, 'n1', 2000)) + tr(tc('a2', W(1000)) + tc('b2', W(2000)))],
    ['of two authors', tr(tc('a1', W(1000)) + ins(11, 'n1', 2000)) + tr(tc('a2', W(1000)) + ins(12, 'n2', 2000, 'w:author="Ann" w:date="2020-01-01T00:00:00Z"'))],
    ['of two dates', tr(tc('a1', W(1000)) + ins(11, 'n1', 2000)) + tr(tc('a2', W(1000)) + ins(12, 'n2', 2000, 'w:author="Zed" w:date="2021-01-01T00:00:00Z"'))],
    ['in other columns', tr(ins(11, 'n1', 1000) + tc('b1', W(2000))) + tr(tc('a2', W(1000)) + ins(12, 'n2', 2000))],
    ['deleted in one row only', tr(tc('a1', W(1000)) + del(11, 'n1', 2000)) + tr(tc('a2', W(1000)) + tc('b2', W(2000)))],
    ['beside a cell over two columns', tr(`<w:tc><w:tcPr>${W(3000)}<w:gridSpan w:val="2"/><w:cellIns w:id="11" ${Z}/></w:tcPr>${p('wide')}</w:tc>`) + tr(tc('a2', W(1000)) + ins(12, 'n2', 2000))],
  ];
  for (const [name, rows] of cases) {
    const o = await open(p('x') + table(rows, [1000, 2000]) + p('y'));
    const before = shape(o.doc);
    const list = revisionsOf(o.doc);
    assert.ok(list.length > 0 && list.every((revision) => revision.kind === 'other' && !revision.canAccept && !revision.canReject), `${name}: ${JSON.stringify(list.map((revision) => revision.kind))}`);
    for (const all of [acceptAll(o.doc), rejectAll(o.doc)]) {
      assert.equal(all.resolved, 0, name);
      assert.ok(all.skipped > 0, name);
      assert.equal(all.ops.length, 0, name);
    }
    assert.throws(() => acceptRevision(o.doc, list[0]), RangeError, name);
    assert.equal(shape(o.doc), before, name);
    checkTables(o.doc, name);
  }
});

test('column: a record of the grid without an inserted column is a change of its own; rejected, the grid it holds comes back when it fits', async () => {
  const rows = tr(tc('a1', W(1200)) + tc('b1', W(1800))) + tr(tc('a2', W(1200)) + tc('b2', W(1800)));
  const fits = await open(p('x') + table(rows, [1200, 1800], 3000, WAS([1000, 2000])) + p('y'));
  assert.deepEqual(summary(fits.doc), ['tableFormat:undefined:0:0']);
  run(fits.history, rejectAll(fits.doc));
  assert.deepEqual([tableOf(fits.doc).grid, tableOf(fits.doc).gridChange], [[1000, 2000], undefined]);
  fits.history.undo();
  run(fits.history, acceptAll(fits.doc));
  assert.deepEqual([tableOf(fits.doc).grid, tableOf(fits.doc).gridChange], [[1200, 1800], undefined]);
  // A record of another number of columns cannot be the grid of these rows: the record goes, the grid stays.
  const other = await open(p('x') + table(rows, [1200, 1800], 3000, WAS([3000])) + p('y'));
  run(other.history, rejectAll(other.doc));
  assert.deepEqual([tableOf(other.doc).grid, tableOf(other.doc).gridChange], [[1200, 1800], undefined]);
  checkTables(other.doc);
});

test('column: the record of the grid somebody made is kept when another column is put in, and goes with the first inserted column', async () => {
  const rows = tr(tc('a1', W(1000)) + ins(11, 'n1', 2000)) + tr(tc('a2', W(1000)) + ins(12, 'n2', 2000));
  const o = await open(p('x') + table(rows, [1000, 2000], 3000, WAS([1000])) + p('y'));
  run(o.history, insertColumn(o.doc, cell(0, 1), 'right', ME));
  assert.deepEqual(tableOf(o.doc).gridChange, { id: 90, grid: [1000] }, 'one record: the first "as it was"');
  assert.deepEqual(summary(o.doc), ['insert:Zed:1:0', 'insert:Me:1:0']);
  // My column rejected: theirs and the record are what they were.
  run(o.history, rejectRevision(o.doc, revisionsOf(o.doc)[1]));
  assert.deepEqual(tableOf(o.doc).gridChange, { id: 90, grid: [1000] });
  assert.deepEqual(summary(o.doc), ['insert:Zed:1:0']);
  // My own column taken back by deleting it leaves their record too.
  const at = run(o.history, insertColumn(o.doc, cell(0, 1), 'right', ME));
  run(o.history, deleteColumn(o.doc, at.focus, ME));
  assert.deepEqual(tableOf(o.doc).gridChange, { id: 90, grid: [1000] });
  assert.deepEqual(summary(o.doc), ['insert:Zed:1:0']);
});

// ------------------------------------------------------------------- commands

test('column: deleting a column of which only some cells are deleted already is refused; one that is all deleted is left as it is', async () => {
  const rows = tr(tc('a1', W(1000)) + del(11, 'n1', 2000)) + tr(tc('a2', W(1000)) + tc('b2', W(2000)));
  const o = await open(p('x') + table(rows, [1000, 2000]) + p('y'));
  const before = shape(o.doc);
  assert.throws(() => deleteColumn(o.doc, cell(1, 1), ME), (error: unknown) => error instanceof RangeError && error.message === TRACK_UNSUPPORTED);
  assert.equal(shape(o.doc), before);
  const whole = await open(FIXED);
  run(whole.history, deleteColumn(whole.doc, cell(0, 1), ME));
  assert.equal(deleteColumn(whole.doc, cell(0, 1), YOU).ops.length, 0, 'nothing is deleted twice');
});

test('column: a row put in without tracking beside a pending column carries the mark of that column, under an id of its own', async () => {
  const rows = tr(tc('a1', W(1000)) + ins(11, 'n1', 2000)) + tr(tc('a2', W(1000)) + ins(12, 'n2', 2000));
  const o = await open(p('x') + table(rows, [1000, 2000], 3000, WAS([1000])) + p('y'), { settings: '' });
  run(o.history, insertRow(o.doc, cell(0, 0), 'below'));
  const xml = await checkFile(o);
  assert.equal((xml.match(/<w:cellIns /g) || []).length, 3, xml);
  assert.deepEqual(summary(o.doc), ['insert:Zed:1:0']);
  run(o.history, rejectAll(o.doc));
  assert.ok(tableOf(o.doc).rows.length === 3 && tableOf(o.doc).rows.every((row) => row.cells.length === 1), 'the column is gone from every row');
  checkTables(o.doc);
});

// --------------------------------------------------- formatting of a row (B3)

test('column: rejecting a change of the formatting of a row leaves the columns the row leaves out as they are', async () => {
  const record = `<w:trPrChange w:id="30" ${Z}><w:trPr><w:cantSplit/></w:trPr></w:trPrChange>`;
  const rows = `<w:tr><w:trPr><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/><w:trHeight w:val="400"/>${record}</w:trPr>${tc('b1', W(2000))}</w:tr>` + tr(tc('a2', W(1000)) + tc('b2', W(2000)) + tc('c2', W(1500)));
  const o = await open(p('x') + table(rows, [1000, 2000, 1500]) + p('y'));
  run(o.history, rejectAll(o.doc));
  const row = tableOf(o.doc).rows[0];
  assert.deepEqual([row.props.gridBefore, row.props.gridAfter, row.props.cantSplit, row.props.height], [1, 1, true, undefined]);
  checkTables(o.doc);
  assert.deepEqual(listRevisions(o.doc).list, []);
});
