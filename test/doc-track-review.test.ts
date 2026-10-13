// Findings of the review of tracked tables, links and pastes: attributes a
// mark of a revision carries beyond its id, author and date; links over
// fields and content controls; and a paragraph mark whose rejection has to
// wait for the table after it.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { writeDocx } from '../src/formats/docx/write';
import type { Table } from '../src/model/doc/types';
import {
  REVISION_UNSUPPORTED,
  TRACK_UNSUPPORTED,
  acceptAll,
  acceptRevision,
  insertHyperlink,
  insertTable,
  rejectAll,
  rejectRevision,
  revisionsOf,
  splitParagraph,
  type DocEditOptions,
} from '../src/ops/doc/index';
import { docPicture, p, pos, r, range, run, tbl, tc, texts, tr } from './doc-ops-fixtures';
import { ME, body, checkFile, open, shape } from './doc-track-fixtures';

const Z = ' w:author="Zed" w:date="2020-01-01T00:00:00Z"';
const NS = 'xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du"';
const UTC = ` ${NS} w16du:dateUtc="2020-01-01T00:00:00Z"`;
const LATER: DocEditOptions = { author: 'Me', date: '2026-10-06T10:00:00Z' };

// ------------------------------------------------- attributes of a mark (B4)

test('review: a cell mark with an attribute of another vocabulary is one change, which is resolved, and the attribute stays with it', async () => {
  for (const kind of ['cellIns', 'cellDel'] as const) {
    const cell = (id: number, text: string) => `<w:tc><w:tcPr><w:${kind} w:id="${id}"${Z}${UTC}/></w:tcPr>${p(text)}</w:tc>`;
    const o = await open(p('before') + tbl(tr(tc('one') + cell(5, 'two')) + tr(tc('three') + cell(6, 'four')), [1500, 1500]) + p('after'));
    const table = o.doc.body.blocks[1] as Table;
    const mark = table.rows[0].cells[1].props[kind === 'cellIns' ? 'inserted' : 'deleted'];
    assert.deepEqual(mark && mark.attrs && mark.attrs.map((pair) => pair[0]), ['xmlns:w16du', 'w16du:dateUtc'], `${kind}: the attribute is part of the mark`);
    assert.equal(table.rows[0].cells[1].props.extras, undefined, `${kind}: no second copy among the kept children`);
    const list = revisionsOf(o.doc);
    assert.deepEqual(list.map((revision) => [revision.kind, revision.canAccept, revision.canReject]), [[kind === 'cellIns' ? 'insert' : 'delete', true, true]], kind);
    // Untouched, the bytes that were read.
    await writeDocx(o.pkg, o.doc);
    assert.deepEqual(await o.pkg.save(), o.bytes, kind);
    for (const verdict of ['accept', 'reject'] as const) {
      const all = verdict === 'accept' ? acceptAll(o.doc) : rejectAll(o.doc);
      assert.deepEqual([all.resolved, all.skipped], [1, 0], `${kind} ${verdict}`);
      run(o.history, all);
      assert.deepEqual(revisionsOf(o.doc), [], `${kind} ${verdict}`);
      const xml = await checkFile(o, `${kind} ${verdict}`);
      assert.equal(xml.indexOf(kind), -1, `${kind} ${verdict}: ${xml}`);
      assert.equal(xml.indexOf('dateUtc'), -1, `${kind} ${verdict}: the attribute went with its mark`);
      o.history.undo();
    }
  }
});

test('review: the mark of a row and of a paragraph keep such an attribute too, once, and write it back', async () => {
  const row = `<w:tr><w:trPr><w:ins w:id="5"${Z}${UTC}/></w:trPr>${tc('three')}</w:tr>`;
  const o = await open(p('before', `<w:rPr><w:del w:id="4"${Z}${UTC}/></w:rPr>`) + p('x') + tbl(tr(tc('one')) + row, [1500]) + p('after'));
  const table = o.doc.body.blocks[2] as Table;
  assert.equal(table.rows[1].props.extras, undefined, 'the row: no second copy');
  assert.deepEqual(table.rows[1].props.inserted && table.rows[1].props.inserted.attrs && table.rows[1].props.inserted.attrs.length, 2);
  const first = o.doc.body.blocks[0];
  assert.ok(first.type === 'paragraph' && first.props.markProps && first.props.markProps.extras === undefined && first.props.markProps.deleted && first.props.markProps.deleted.attrs, 'the paragraph mark: no second copy');
  assert.deepEqual(revisionsOf(o.doc).map((revision) => revision.kind), ['delete', 'insert']);
  // Written from the model (another edit of the same table and paragraph), the marks carry what they carried.
  run(o.history, splitParagraph(o.doc, pos(0, 2)));
  run(o.history, insertTable(o.doc, { story: 'body', path: [3, 0, 0, 0], offset: 0 }, 1, 1));
  const xml = await checkFile(o);
  assert.equal(xml.split('w16du:dateUtc="2020-01-01T00:00:00Z"').length - 1, 2, xml);
  assert.ok(xml.indexOf(`<w:ins w:id="5"${Z}${UTC}/>`) > 0, xml);
  const all = rejectAll(o.doc);
  assert.equal(all.skipped, 0);
  run(o.history, all);
  assert.equal((await body(o)).indexOf('dateUtc'), -1);
});

test('review: a mark that holds an element is not one the model can stand for: it is kept as it is', async () => {
  const odd = `<w:tc><w:tcPr><w:cellIns w:id="5"${Z}><w:odd/></w:cellIns></w:tcPr>${p('two')}</w:tc>`;
  const o = await open(p('x') + tbl(tr(tc('one') + odd), [1500, 1500]) + p('y'));
  const cell = (o.doc.body.blocks[1] as Table).rows[0].cells[1];
  assert.equal(cell.props.inserted, undefined);
  assert.equal(cell.props.extras && cell.props.extras.length, 1);
  assert.deepEqual(revisionsOf(o.doc).map((revision) => revision.kind), ['other']);
});

// ------------------------------------------------------------------ links (B5)

const LINK = 'https://example.org/';
const refused = (make: () => unknown): void => {
  assert.throws(make, (error: unknown) => error instanceof RangeError && error.message === TRACK_UNSUPPORTED);
};

test('review: a tracked link over, or into, a simple field or a content control is refused, and nothing is written twice', async () => {
  const field = p(r('page ') + '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>' + r(' of'));
  const control = p(r('a ') + `<w:sdt><w:sdtPr><w:id w:val="7"/></w:sdtPr><w:sdtContent>${r('inside')}</w:sdtContent></w:sdt>` + r(' b'));
  const o = await open(field + control);
  const before = docPicture(o.doc);
  const relationships = o.doc.body.relationships;
  refused(() => insertHyperlink(o.doc, range(pos(0, 3), pos(0, 8)), LINK, ME));
  refused(() => insertHyperlink(o.doc, range(pos(0, 5), pos(0, 6)), LINK, ME));
  refused(() => insertHyperlink(o.doc, range(pos(1, 3), pos(1, 6)), LINK, ME));
  refused(() => insertHyperlink(o.doc, range(pos(1, 1), pos(1, 9)), LINK, ME));
  assert.equal(docPicture(o.doc), before);
  assert.ok(o.doc.body.relationships === relationships, 'no relationship is left behind');
  // Beside them it is a link like any other.
  run(o.history, insertHyperlink(o.doc, range(pos(0, 0), pos(0, 4)), LINK, ME));
  run(o.history, insertHyperlink(o.doc, range(pos(1, 9), pos(1, 10)), LINK, ME));
  const xml = await checkFile(o);
  assert.equal(xml.split('<w:fldSimple ').length - 1, 1, xml);
  assert.equal(xml.split('<w:sdt>').length - 1, 1, xml);
});

// ----------------------------------------- a mark that waits for a table (B6)

test('review: the mark before an inserted table cannot be rejected before the table is; then every order gives what "reject all" gives', async () => {
  const make = async () => {
    const o = await open(p('abcd', '<w:jc w:val="center"/>') + p('end'));
    run(o.history, splitParagraph(o.doc, pos(0, 2), ME));
    run(o.history, insertTable(o.doc, pos(1, 0), 1, 1, LATER));
    return o;
  };
  const all = await make();
  const start = shape((await open(p('abcd', '<w:jc w:val="center"/>') + p('end'))).doc);
  run(all.history, rejectAll(all.doc));
  assert.equal(shape(all.doc), start, 'all rejected: the paragraph is one again');

  const o = await make();
  const list = revisionsOf(o.doc);
  assert.deepEqual(list.map((revision) => `${revision.kind}:${revision.date}:${revision.rows}`), ['insert:2026-10-05T10:00:00Z:0', 'insert:2026-10-06T10:00:00Z:1']);
  const before = docPicture(o.doc);
  // Alone, the mark would stay as a plain mark, and the paragraph cut in two for good.
  assert.throws(() => rejectRevision(o.doc, list[0]), (error: unknown) => error instanceof RangeError && error.message === REVISION_UNSUPPORTED);
  assert.equal(docPicture(o.doc), before);
  // Accepting it is nothing that has to wait.
  run(o.history, acceptRevision(o.doc, list[0]));
  o.history.undo();
  run(o.history, rejectRevision(o.doc, revisionsOf(o.doc)[1]));
  run(o.history, rejectRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.equal(shape(o.doc), start, 'the table first, then the mark: the same');
  assert.deepEqual(texts(o.doc), ['abcd', 'end']);
});

test('review: the deleted mark before a deleted table cannot be accepted before the table is', async () => {
  const gone = (text: string) => `<w:tr><w:trPr><w:del w:id="21" w:author="Zed" w:date="2021-01-01T00:00:00Z"/></w:trPr>${tc(p(`<w:del w:id="22" w:author="Zed" w:date="2021-01-01T00:00:00Z"><w:r><w:delText>${text}</w:delText></w:r></w:del>`))}</w:tr>`;
  const make = () => open(p('ab', `<w:jc w:val="center"/><w:rPr><w:del w:id="20"${Z}/></w:rPr>`) + tbl(gone('cell'), [1500]) + p('cd') + p('end'));
  const all = await make();
  run(all.history, acceptAll(all.doc));
  assert.deepEqual(texts(all.doc), ['abcd', 'end']);
  const want = shape(all.doc);

  const o = await make();
  const list = revisionsOf(o.doc);
  assert.equal(list.length, 2);
  assert.throws(() => acceptRevision(o.doc, list[0]), (error: unknown) => error instanceof RangeError && error.message === REVISION_UNSUPPORTED);
  // Rejecting it takes nothing out: that does not wait.
  run(o.history, rejectRevision(o.doc, list[0]));
  o.history.undo();
  run(o.history, acceptRevision(o.doc, revisionsOf(o.doc)[1]));
  run(o.history, acceptRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.equal(shape(o.doc), want);
});

test('review: the deleted mark before a table whose every column is deleted waits for the columns too', async () => {
  const gone = (id: number, text: string) => `<w:tc><w:tcPr><w:cellDel w:id="${id}" w:author="Zed" w:date="2021-01-01T00:00:00Z"/></w:tcPr>${p(text)}</w:tc>`;
  const make = () => open(p('ab', `<w:rPr><w:del w:id="20"${Z}/></w:rPr>`) + tbl(tr(gone(21, 'c1')) + tr(gone(22, 'c2')), [1500]) + p('cd') + p('end'));
  const all = await make();
  run(all.history, acceptAll(all.doc));
  assert.deepEqual(texts(all.doc), ['abcd', 'end']);
  const o = await make();
  const list = revisionsOf(o.doc);
  assert.deepEqual(list.map((revision) => `${revision.kind}:${revision.columns}`), ['delete:0', 'delete:1']);
  assert.throws(() => acceptRevision(o.doc, list[0]), (error: unknown) => error instanceof RangeError && error.message === REVISION_UNSUPPORTED);
  run(o.history, acceptRevision(o.doc, list[1]));
  run(o.history, acceptRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.equal(shape(o.doc), shape(all.doc));
});

test('review: the empty paragraph after an inserted table at the end of a cell waits for the table', async () => {
  const make = async () => {
    const o = await open(p('x') + tbl(tr(tc('cell')), [3000]) + p('end'));
    // At the end of the cell: the table, and the paragraph a cell has to end with.
    run(o.history, insertTable(o.doc, { story: 'body', path: [1, 0, 0, 0], offset: 4 }, 1, 1, ME));
    return o;
  };
  const start = shape((await open(p('x') + tbl(tr(tc('cell')), [3000]) + p('end'))).doc);
  // Two changes: the paragraph is made at another moment than the table.
  const o = await open(
    p('x') +
      tbl(
        tr(
          tc(
            p('cell') +
              tbl(`<w:tr><w:trPr><w:ins w:id="30" w:author="Zed" w:date="2021-01-01T00:00:00Z"/></w:trPr>${tc('in')}</w:tr>`, [1000]) +
              p('', `<w:rPr><w:ins w:id="31"${Z}/></w:rPr>`),
          ),
        ),
        [3000],
      ) +
      p('end'),
  );
  const list = revisionsOf(o.doc);
  assert.deepEqual(list.map((revision) => `${revision.kind}:${revision.rows}:${revision.marks}`), ['insert:1:0', 'insert:0:1']);
  assert.throws(() => rejectRevision(o.doc, list[1]), (error: unknown) => error instanceof RangeError && error.message === REVISION_UNSUPPORTED);
  run(o.history, rejectRevision(o.doc, list[0]));
  run(o.history, rejectRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.equal(shape(o.doc), start);
  // Made by one command they are one change, and go together.
  const one = await make();
  assert.equal(revisionsOf(one.doc).length, 1);
  run(one.history, rejectRevision(one.doc, revisionsOf(one.doc)[0]));
  assert.equal(shape(one.doc), start);
});

test('review: an empty paragraph whose mark has to go waits for the paragraph before it whose mark has to go too', async () => {
  const mark = (id: number, date: string) => `<w:rPr><w:ins w:id="${id}" w:author="Zed" w:date="${date}"/></w:rPr>`;
  const make = () => open(p('x') + tbl(tr(tc('a')), [1000]) + p('', '<w:jc w:val="right"/>' + mark(40, '2021-01-01T00:00:00Z')) + p('', mark(41, '2022-01-01T00:00:00Z')) + tbl(tr(tc('b')), [1000]) + p('end'));
  const all = await make();
  run(all.history, rejectAll(all.doc));
  const want = shape(all.doc);
  const o = await make();
  const list = revisionsOf(o.doc);
  assert.equal(list.length, 2);
  // The second alone would decide what stays between the tables before the first has joined it.
  assert.throws(() => rejectRevision(o.doc, list[1]), (error: unknown) => error instanceof RangeError && error.message === REVISION_UNSUPPORTED);
  run(o.history, rejectRevision(o.doc, list[0]));
  run(o.history, rejectRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.equal(shape(o.doc), want);
  assert.deepEqual(revisionsOf(o.doc), []);
});

test('review: a deleted mark whose paragraph holds nothing but deleted text waits for the text: the paragraph goes with the last of them', async () => {
  const first = p(`<w:del w:id="50" w:author="Zed" w:date="2021-01-01T00:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del>`, `<w:rPr><w:del w:id="51"${Z}/></w:rPr>`);
  const make = () => open(p('x') + tbl(tr(tc(first + tbl(tr(tc('in')), [1000]) + p('last'))), [3000]) + p('end'));
  const all = await make();
  run(all.history, acceptAll(all.doc));
  assert.deepEqual(texts(all.doc), ['x', 'in', 'last', 'end']);
  const o = await make();
  const mark = revisionsOf(o.doc).find((revision) => revision.marks === 1);
  assert.ok(mark, 'the deleted mark is a change of its own');
  // Alone it would stay, a plain mark of a paragraph that is empty once the text is accepted too.
  assert.throws(() => acceptRevision(o.doc, mark), (error: unknown) => error instanceof RangeError && error.message === REVISION_UNSUPPORTED);
  run(o.history, acceptRevision(o.doc, revisionsOf(o.doc)[0]));
  run(o.history, acceptRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.equal(shape(o.doc), shape(all.doc));
  assert.deepEqual(revisionsOf(o.doc), []);
});

test('review: a change of which only a mark has to wait is resolved but for that mark, which stays a change', async () => {
  // One change: the deleted mark, the deleted row of the table after it. The other row is deleted by somebody else.
  const row = (id: number, who: string, text: string) => `<w:tr><w:trPr><w:del w:id="${id}" ${who}/></w:trPr>${tc(text)}</w:tr>`;
  const zed = 'w:author="Zed" w:date="2020-01-01T00:00:00Z"';
  const o = await open(p('ab', `<w:rPr><w:del w:id="60" ${zed}/></w:rPr>`) + tbl(row(61, zed, 'r1') + row(62, 'w:author="Ann" w:date="2020-01-01T00:00:00Z"', 'r2'), [1000]) + p('cd'));
  const list = revisionsOf(o.doc);
  assert.deepEqual(list.map((revision) => `${revision.author}:${revision.marks}:${revision.rows}`), ['Zed:1:1', 'Ann:0:1']);
  run(o.history, acceptRevision(o.doc, list[0]));
  // The row is gone; the mark waits for the table, which would go with Ann's row.
  assert.deepEqual(revisionsOf(o.doc).map((revision) => `${revision.author}:${revision.marks}:${revision.rows}`), ['Zed:1:0', 'Ann:0:1']);
  assert.deepEqual(texts(o.doc), ['ab', 'r2', 'cd']);
  run(o.history, acceptRevision(o.doc, revisionsOf(o.doc)[1]));
  run(o.history, acceptRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.deepEqual(texts(o.doc), ['abcd']);
  assert.deepEqual(revisionsOf(o.doc), []);
});

test('review: a mark before a table that stays is still resolved as before: it stays, a plain mark', async () => {
  const o = await open(p('ab', `<w:rPr><w:ins w:id="20"${Z}/></w:rPr>`) + tbl(tr(tc('cell')), [1500]) + p('end'));
  run(o.history, rejectAll(o.doc));
  assert.deepEqual(texts(o.doc), ['ab', 'cell', 'end']);
  assert.deepEqual(revisionsOf(o.doc), []);
});
