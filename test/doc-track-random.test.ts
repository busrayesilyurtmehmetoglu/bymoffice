// Random tracked edits of every kind (seeded): typing, deleting, formatting,
// rows, columns, tables, pictures, links, pasted content and blocks. What has
// to hold whatever the order: every file written on the way is well formed,
// in schema order and without an id twice; read again it is the model that
// was written; with one author, rejecting everything gives the text, the
// formatting and the tables of the start, and accepting everything leaves no
// change; undoing everything gives the bytes that were read; and a change
// resolved by itself takes no change of somebody else with it.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { writeDocx } from '../src/formats/docx/write';
import type { Block, DocModel } from '../src/model/doc/types';
import {
  acceptAll,
  acceptRevision,
  deleteBackward,
  deleteColumn,
  deleteForward,
  deleteRange,
  deleteRow,
  fragmentOf,
  insertBlock,
  insertColumn,
  insertFragment,
  insertHyperlink,
  insertImage,
  insertPageBreak,
  insertRow,
  insertTable,
  insertText,
  listRevisions,
  paragraphLength,
  rejectAll,
  rejectRevision,
  revisionsOf,
  setParaFormat,
  setRunFormat,
  splitParagraph,
  storyIds,
  storyOf,
  type DocEdit,
  type DocEditOptions,
  type DocFragment,
  type DocPosition,
} from '../src/ops/doc/index';
import { checkInvariants, dice, p, paragraphAt, paragraphPaths, picture, pos, r, randomPosition, range, run, tbl, tc, tr } from './doc-ops-fixtures';
import { ME, YOU, body, changes, checkFile, checkReopened, checkTables, open, revisionIds, shape, written, type Opened } from './doc-track-fixtures';

type Dice = ReturnType<typeof dice>;

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4]);

const BODY =
  p('alpha beta gamma') +
  p(r('bold', '<w:b/>') + r(' and plain') + r(' italic', '<w:i/>'), '<w:jc w:val="right"/>') +
  p('') +
  tbl(tr(tc('one') + tc(p('two') + p(r('three', '<w:u w:val="single"/>')))) + tr(tc('four') + tc('five')), [1500, 1500]) +
  p('between') +
  // A table of a fixed width: it has to stay as wide as its grid whatever is resolved, in whatever order.
  `<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="2000"/></w:tblGrid>${tr(tc('f1', '<w:tcW w:w="1000" w:type="dxa"/>') + tc('f2', '<w:tcW w:w="2000" w:type="dxa"/>')) + tr(tc('f3', '<w:tcW w:w="1000" w:type="dxa"/>') + tc('f4', '<w:tcW w:w="2000" w:type="dxa"/>'))}</w:tbl>` +
  p('omega end', '<w:pStyle w:val="Heading1"/><w:spacing w:before="120"/>');

const SOURCE =
  p(r('plain ') + r('bold', '<w:b/>') + r(' text'), '<w:jc w:val="center"/>') + tbl(tr(tc('c1') + tc('c2')), [1200, 1200]) + p(r('see ') + `<w:hyperlink r:id="rIdL">${r('link')}</w:hyperlink>` + r(' end'));

/** The fragments a paste takes: text of one paragraph, paragraphs, and paragraphs with a table and a link. */
async function fragments(): Promise<DocFragment[]> {
  const source = await open(SOURCE, { rels: [['rIdL', 'hyperlink', 'https://example.org/', true]] });
  return [fragmentOf(source.doc, range(pos(0, 3), pos(0, 12))), fragmentOf(source.doc, range(pos(0, 2), pos(2, 3))), fragmentOf(source.doc, range(pos(0, 0), pos(2, 12))), fragmentOf(source.doc, range(pos(2, 2), pos(2, 9)))];
}

/**
 * The text, the formatting and the tables of a document, run boundaries and
 * the attributes of elements aside: what rejecting everything has to bring
 * back. Tracked changes are part of it, so that a document with any left is
 * another document.
 */
function structure(doc: DocModel): string {
  const blocks = (list: Block[], depth: number): unknown[] => {
    const out: unknown[] = [];
    if (depth > 32) return out;
    for (const block of list) {
      if (block.type === 'paragraph') {
        const spans: [string, string][] = [];
        for (const inline of block.children) {
          if (inline.type !== 'run') continue;
          let text = '';
          for (const item of inline.content) text += item.type === 'text' || item.type === 'deletedText' ? item.text : `<${item.type}>`;
          if (text === '') continue;
          const props = picture([inline.props, inline.wrap]);
          const last = spans[spans.length - 1];
          if (last && last[1] === props) last[0] += text;
          else spans.push([text, props]);
        }
        out.push(['p', picture(block.props), spans]);
      } else if (block.type === 'table') {
        out.push([
          'table',
          picture(block.props),
          block.grid,
          picture(block.gridChange),
          block.rows.map((row) => [picture(row.props), row.cells.map((cell) => [picture(cell.props), blocks(cell.blocks, depth + 1)])]),
        ]);
      } else out.push([block.type]);
    }
    return out;
  };
  const out: unknown[] = [];
  for (const story of storyIds(doc)) out.push([story, blocks(storyOf(doc, story).blocks, 0)]);
  return JSON.stringify(out);
}

/**
 * Resolves the changes of a document one by one in a random order, each with the same verdict, until none is left.
 * A change that has to wait for another (the mark before an inserted table) is passed over for the moment.
 */
function oneByOne(o: Opened, d: Dice, verdict: 'accept' | 'reject', what: string): void {
  for (let guard = 0; guard < 400; guard++) {
    const list = revisionsOf(o.doc);
    if (!list.length) return;
    let from = d.int(list.length);
    let edit: DocEdit | undefined;
    for (let tries = 0; tries < list.length && !edit; tries++, from++) {
      const revision = list[from % list.length];
      try {
        edit = verdict === 'accept' ? acceptRevision(o.doc, revision) : rejectRevision(o.doc, revision);
      } catch (error) {
        assert.ok(error instanceof RangeError, `${what}: ${String(error)}`);
      }
    }
    assert.ok(edit && edit.ops.length > 0, `${what}: ${list.length} changes are left and none of them can be resolved`);
    run(o.history, edit);
    checkTables(o.doc, what);
  }
  assert.fail(`${what}: the list of changes does not get shorter`);
}

/** A place in a table cell of the body, when there is one. */
function randomCell(doc: DocModel, d: Dice): DocPosition | undefined {
  const inside = paragraphPaths(doc, 'body').filter((path) => path.length >= 4);
  return inside.length ? { story: 'body', path: d.pick(inside), offset: 0 } : undefined;
}

const RUN_PATCHES = [{ bold: true }, { bold: null }, { italic: true }, { size: 30 }, { color: 'FF0000' }, { style: 'Strong' }, { style: null }];
const PARA_PATCHES = [{ align: 'center' }, { align: null }, { indent: { left: 400 } }, { style: 'Quote' }, { style: null }];

function randomEdit(o: Opened, d: Dice, who: DocEditOptions, pastes: DocFragment[]): DocEdit | 'undo' | undefined {
  const doc = o.doc;
  const story = d.chance(0.85) ? 'body' : d.pick(storyIds(doc));
  const paths = paragraphPaths(doc, story);
  const at = randomPosition(doc, d, story, paths);
  const other = randomPosition(doc, d, story, paths);
  const cell = randomCell(doc, d);
  switch (d.int(20)) {
    case 0:
    case 1:
      return insertText(doc, at, d.pick(['x', 'two words', 'a\nb', '\t', ' ']), undefined, who);
    case 2:
      return deleteBackward(doc, at, 'char', who);
    case 3:
      return deleteForward(doc, at, d.chance(0.3) ? 'word' : 'char', who);
    case 4:
      return deleteRange(doc, range(at, other), who);
    case 5:
      return splitParagraph(doc, at, who);
    case 6:
      return setRunFormat(doc, range(at, other), d.pick(RUN_PATCHES), who);
    case 7:
      return setParaFormat(doc, at, d.pick(PARA_PATCHES), who);
    case 8:
      return cell ? insertRow(doc, cell, d.chance(0.5) ? 'above' : 'below', who) : undefined;
    case 9:
      return cell ? deleteRow(doc, cell, who) : undefined;
    case 10:
      return cell ? insertColumn(doc, cell, d.chance(0.5) ? 'left' : 'right', who) : undefined;
    case 11:
      return cell ? deleteColumn(doc, cell, who) : undefined;
    case 12:
      return insertTable(doc, at, 1 + d.int(2), 1 + d.int(2), who);
    case 13:
      return insertImage(doc, at, { bytes: PNG, type: 'image/png', width: 95250, height: 95250 }, who);
    case 14: {
      // A link lies in one paragraph: over the next few characters, or at the caret.
      const end = Math.min(paragraphLength(paragraphAt(doc, at.path, story)), at.offset + d.int(5));
      return insertHyperlink(doc, range(at, { ...at, offset: end }), 'https://example.org/' + d.int(3), who);
    }
    case 15:
    case 16:
      return insertFragment(doc, d.chance(0.8) ? at : range(at, other), d.pick(pastes), who);
    case 17:
      return insertPageBreak(doc, at, who);
    case 18:
      return insertBlock(doc, at, { type: 'paragraph', props: {}, children: [{ type: 'run', props: {}, content: [{ type: 'text', text: 'block' }] }] }, who);
    default:
      return 'undo';
  }
}

test('random tracked edits of every kind by one author: files are sound, rejecting all gives the document back, accepting all leaves no change', async () => {
  const pastes = await fragments();
  const kinds = new Set<string>();
  for (let seed = 1; seed <= 24; seed++) {
    const d = dice(seed * 4409);
    const o = await open(BODY, { headers: { rIdH: p('header text') } });
    const before = structure(o.doc);
    let made = 0;
    for (let step = 0; step < 40; step++) {
      let edit: DocEdit | 'undo' | undefined;
      try {
        edit = randomEdit(o, d, ME, pastes);
      } catch (error) {
        // What cannot be tracked, a place that cannot be edited: refused, and nothing happened.
        assert.ok(error instanceof RangeError, `seed ${seed} step ${step}: ${String(error)}`);
        continue;
      }
      if (edit === 'undo') {
        if (o.history.canUndo && d.chance(0.5)) o.history.undo();
        continue;
      }
      if (!edit || edit.ops.length === 0) continue;
      run(o.history, edit);
      made++;
      checkInvariants(o.doc, false);
      checkTables(o.doc, `seed ${seed} step ${step}`);
      if (step % 8 === 0) {
        await checkFile(o, `seed ${seed} step ${step}`);
        await checkReopened(o, `seed ${seed} step ${step}`, false);
      }
    }
    assert.ok(made > 8, `seed ${seed}: ${made} edits`);
    await checkFile(o, `seed ${seed}`);
    const again = await checkReopened(o, `seed ${seed}`, false);
    assert.ok(structure(again) === structure(o.doc), `seed ${seed}: read again, the text, the formatting or a table is another`);
    for (const revision of revisionsOf(o.doc)) kinds.add(revision.columns > 0 ? 'column' : revision.kind);

    // Everything rejected: the document of the start.
    const rejected = rejectAll(o.doc);
    assert.equal(rejected.skipped, 0, `seed ${seed}: changes passed over by "reject all"`);
    run(o.history, rejected);
    assert.equal(changes(o.doc), '', `seed ${seed}: changes left after "reject all"`);
    assert.ok(structure(o.doc) === before, `seed ${seed}: rejecting everything does not give the document back`);
    checkInvariants(o.doc, false);
    checkTables(o.doc, `seed ${seed} rejected`);
    const plain = await checkFile(o, `seed ${seed} rejected`);
    assert.deepEqual(revisionIds(plain), [], `seed ${seed}: rejected`);
    const rejectedBody = await body(o);
    if (rejected.ops.length) o.history.undo();

    // Every change rejected by itself, in a random order: the same document, model and file.
    const depth = undoDepth(o);
    oneByOne(o, d, 'reject', `seed ${seed}, rejecting one by one`);
    assert.ok(structure(o.doc) === before, `seed ${seed}: rejecting one by one does not give the document back`);
    assert.equal(await body(o), rejectedBody, `seed ${seed}: rejecting one by one writes another file than "reject all"`);
    undoTo(o, depth);

    // Everything accepted: no change is left, and the file says nothing of any.
    const accepted = acceptAll(o.doc);
    assert.equal(accepted.skipped, 0, `seed ${seed}: changes passed over by "accept all"`);
    run(o.history, accepted);
    assert.equal(changes(o.doc), '', `seed ${seed}: changes left after "accept all"`);
    checkInvariants(o.doc, false);
    checkTables(o.doc, `seed ${seed} accepted`);
    const clean = await checkFile(o, `seed ${seed} accepted`);
    assert.deepEqual(revisionIds(clean), [], `seed ${seed}: accepted`);
    await checkReopened(o, `seed ${seed} accepted`, false);
    const acceptedShape = shape(o.doc);
    const acceptedBody = await body(o);
    if (accepted.ops.length) o.history.undo();
    // And accepted by itself, in a random order.
    const before2 = undoDepth(o);
    oneByOne(o, d, 'accept', `seed ${seed}, accepting one by one`);
    assert.ok(shape(o.doc) === acceptedShape, `seed ${seed}: accepting one by one leaves another document than "accept all"`);
    assert.equal(await body(o), acceptedBody, `seed ${seed}: accepting one by one writes another file than "accept all"`);
    undoTo(o, before2);

    // And undoing all of it gives the file back.
    while (o.history.canUndo) o.history.undo();
    await writeDocx(o.pkg, o.doc);
    assert.deepEqual(await o.pkg.save(), o.bytes, `seed ${seed}: bytes`);
  }
  // The runs together made changes of every kind the commands make.
  for (const kind of ['insert', 'delete', 'format', 'paragraphFormat', 'column']) assert.ok(kinds.has(kind), `no change of kind ${kind} was ever made`);
});

/** How many steps of the history can be undone, and undoing back to such a count. */
function undoDepth(o: Opened): number {
  let depth = 0;
  while (o.history.canUndo) {
    o.history.undo();
    depth++;
  }
  for (let i = 0; i < depth; i++) o.history.redo();
  return depth;
}

function undoTo(o: Opened, depth: number): void {
  let now = undoDepth(o);
  for (; now > depth; now--) o.history.undo();
}

/** How many elements of tracked changes in the XML name an author. */
const signedBy = (xml: string, author: string): number => xml.split(` w:author="${author}"`).length - 1;

test('random tracked edits by two authors, with changes resolved one by one: a change that takes no content out takes no change of the other author with it', async () => {
  const pastes = await fragments();
  let alone = 0;
  for (let seed = 1; seed <= 20; seed++) {
    const d = dice(seed * 7577);
    const theirs = `<w:tr><w:trPr><w:ins w:id="60" w:author="Zed" w:date="2020-01-01T00:00:00Z"/></w:trPr>${tc('theirs')}<w:tc><w:tcPr><w:cellDel w:id="61" w:author="Zed" w:date="2020-01-01T00:00:00Z"/></w:tcPr>${p('gone')}</w:tc></w:tr>`;
    const o = await open(BODY + tbl(theirs, [1500, 1500]) + p('last'));
    let made = 0;
    for (let step = 0; step < 45; step++) {
      const who = d.chance(0.6) ? ME : YOU;
      let edit: DocEdit | 'undo' | undefined;
      let resolved: { author?: string; keeps: boolean } | undefined;
      try {
        if (d.chance(0.2)) {
          const accept = d.chance(0.5);
          const list = listRevisions(o.doc).list.filter((revision) => (accept ? revision.canAccept : revision.canReject));
          if (!list.length) continue;
          const revision = d.pick(list);
          // Accepting an insertion and rejecting a deletion take nothing out; neither does a verdict on formatting.
          const keeps = revision.kind === 'insert' ? accept : revision.kind === 'delete' ? !accept : revision.kind !== 'moveFrom' && revision.kind !== 'moveTo';
          resolved = { author: revision.author, keeps };
          edit = accept ? acceptRevision(o.doc, revision) : rejectRevision(o.doc, revision);
        } else edit = randomEdit(o, d, who, pastes);
      } catch (error) {
        assert.ok(error instanceof RangeError, `seed ${seed} step ${step}: ${String(error)}`);
        continue;
      }
      if (edit === 'undo') {
        if (o.history.canUndo && d.chance(0.5)) o.history.undo();
        continue;
      }
      if (!edit || edit.ops.length === 0) continue;
      const before = resolved && resolved.keeps ? await written(o) : undefined;
      run(o.history, edit);
      made++;
      checkInvariants(o.doc, false);
      checkTables(o.doc, `seed ${seed} step ${step}`);
      if (before !== undefined && resolved) {
        const after = await written(o);
        for (const author of ['Me', 'You', 'Zed']) {
          if (author === resolved.author) continue;
          assert.equal(signedBy(after, author), signedBy(before, author), `seed ${seed} step ${step}: a change of ${String(resolved.author)} resolved alone took changes of ${author} with it`);
        }
        alone++;
      }
      if (step % 9 === 0) await checkFile(o, `seed ${seed} step ${step}`);
    }
    assert.ok(made > 8, `seed ${seed}: ${made} edits`);
    await checkFile(o, `seed ${seed}`);
    await checkReopened(o, `seed ${seed}`, false);
    const accepted = acceptAll(o.doc);
    assert.equal(accepted.skipped, 0, `seed ${seed}: changes passed over by "accept all"`);
    run(o.history, accepted);
    assert.equal(changes(o.doc), '', `seed ${seed}: changes left after "accept all"`);
    const clean = await checkFile(o, `seed ${seed} accepted`);
    assert.deepEqual(revisionIds(clean), [], `seed ${seed}: accepted`);
    while (o.history.canUndo) o.history.undo();
    await writeDocx(o.pkg, o.doc);
    assert.deepEqual(await o.pkg.save(), o.bytes, `seed ${seed}: bytes`);
  }
  assert.ok(alone > 20, `${alone} changes were resolved alone`);
});
