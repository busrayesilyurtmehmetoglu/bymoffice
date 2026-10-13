// Tracked changes as a list, and accepting and rejecting them: one at a time
// and all at once. What is accepted or rejected changes nothing but itself,
// undoing gives back the bytes that were read, and what the engine cannot
// resolve is refused or skipped and counted, never half done.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { buildDocx as convertToDocx, writeDocx } from '../src/formats/docx/write';
import type { DocModel, Paragraph, Table } from '../src/model/doc/types';
import {
  DocHistory,
  REVISIONS_PROTECTED,
  REVISION_GONE,
  REVISION_UNSUPPORTED,
  acceptAll,
  acceptRevision,
  deleteRange,
  insertText,
  listRevisions,
  rejectAll,
  rejectRevision,
  revisionsOf,
  type DocRange,
  type DocRevision,
} from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { parseXml } from '../src/xml';
import { DocBuilder, FC, sprm, sttb, u16le } from './doc-binary-fixtures';
import { STYLES, checkInvariants, p, pos, r, range, run, tbl, tc, texts, tr } from './doc-ops-fixtures';
import { buildDocx as buildSpec, type DocxSpec } from './docx-fixtures';
import { schemaOrderViolations } from './docx-write-helpers.test';
import { read as readRtf } from './rtf-fixtures';

const decoder = new TextDecoder();
const SECT = '<w:sectPr><w:pgSz w:w="9000" w:h="9000"/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr>';

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

async function written(o: Opened, part = '/word/document.xml'): Promise<string> {
  await writeDocx(o.pkg, o.doc);
  const text = decoder.decode(await o.pkg.readPart(part));
  parseXml(text);
  assert.deepEqual(schemaOrderViolations(text), []);
  return text;
}

/** The body of the written document part, without the last section. */
async function body(o: Opened): Promise<string> {
  const text = await written(o);
  return text.slice(text.indexOf('<w:body>') + 8, text.lastIndexOf('<w:sectPr>'));
}

const WHO = (author: string, date = '2026-10-01T09:00:00Z') => ` w:author="${author}" w:date="${date}"`;
const ins = (id: number, inner: string, author = 'Ann', date?: string) => `<w:ins w:id="${id}"${WHO(author, date)}>${inner}</w:ins>`;
const del = (id: number, inner: string, author = 'Ann', date?: string) => `<w:del w:id="${id}"${WHO(author, date)}>${inner}</w:del>`;
/** As the writer writes text: outer spaces ask for xml:space. */
const space = (text: string) => (/^\s|\s$/.test(text) ? ' xml:space="preserve"' : '');
const t = (text: string) => `<w:r><w:t${space(text)}>${text}</w:t></w:r>`;
const dt = (text: string) => `<w:r><w:delText${space(text)}>${text}</w:delText></w:r>`;
const markIns = (id: number, author = 'Ann') => `<w:rPr><w:ins w:id="${id}"${WHO(author)}/></w:rPr>`;
const markDel = (id: number, author = 'Ann') => `<w:rPr><w:del w:id="${id}"${WHO(author)}/></w:rPr>`;

const kinds = (doc: DocModel): string[] => revisionsOf(doc).map((revision) => revision.kind);
const REVISION_TAGS = /<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|sectPrChange|tblPrChange|trPrChange|tcPrChange|moveFromRangeStart|moveToRangeStart|moveFromRangeEnd|moveToRangeEnd|delText|delInstrText)\b/;

/** Applies a command, checks the document, and that undoing it gives the file back byte for byte. */
async function applied(o: Opened, edit: { ops: unknown[]; selection: DocRange }): Promise<string> {
  run(o.history, edit as never);
  checkInvariants(o.doc);
  const out = await body(o);
  const saved = await o.pkg.save();
  assert.ok(saved.length > 0);
  o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes, 'undo gives the bytes that were read');
  o.history.redo();
  assert.equal(await body(o), out, 'redo gives the same document again');
  return out;
}

const first = (doc: DocModel, kind?: string): DocRevision => {
  const found = revisionsOf(doc).find((revision) => kind === undefined || revision.kind === kind);
  assert.ok(found, `no revision${kind ? ' of kind ' + kind : ''}`);
  return found;
};

const refused = (make: () => unknown, message: string): void => {
  assert.throws(make, (error: unknown) => error instanceof RangeError && error.message === message);
};

// ----------------------------------------------------------------------- list

test('revisions: the list names kind, author, date, place and text of every tracked change', async () => {
  const o = await open(p(t('one ') + ins(1, t('two ')) + del(2, dt('three '), 'Bob', '2026-10-02T10:00:00Z') + t('four')));
  const list = revisionsOf(o.doc);
  assert.equal(list.length, 2);
  assert.deepEqual(
    list.map((revision) => [revision.index, revision.kind, revision.author, revision.date, revision.text, revision.story]),
    [
      [0, 'insert', 'Ann', '2026-10-01T09:00:00Z', 'two ', 'body'],
      [1, 'delete', 'Bob', '2026-10-02T10:00:00Z', 'three ', 'body'],
    ],
  );
  assert.deepEqual([list[0].start, list[0].end], [pos(0, 4), pos(0, 8)]);
  // Deleted text takes no offsets: the deletion is a place, not a stretch.
  assert.deepEqual([list[1].start, list[1].end], [pos(0, 8), pos(0, 8)]);
  assert.ok(list.every((revision) => revision.canAccept && revision.canReject));
});

test('revisions: neighbouring pieces of one author and date are one change; another author, another kind or text in between starts a new one', async () => {
  const o = await open(
    p(ins(1, t('a')) + ins(2, t('b')) + ins(3, t('c'), 'Bob') + del(4, dt('d'), 'Bob') + ins(5, t('e'), 'Bob') + t(' plain ') + ins(6, t('f'), 'Bob') + ins(7, t('g'), 'Bob', '2026-10-03T00:00:00Z')),
  );
  assert.deepEqual(
    revisionsOf(o.doc).map((revision) => `${revision.kind}:${revision.author}:${revision.text}`),
    ['insert:Ann:ab', 'insert:Bob:c', 'delete:Bob:d', 'insert:Bob:e', 'insert:Bob:f', 'insert:Bob:g'],
  );
});

test('revisions: an insertion that runs over a paragraph mark is one change, and so is an inserted table', async () => {
  const o = await open(
    p(t('keep ') + ins(1, t('new')), markIns(2)) +
      p(ins(3, t('more')) + t(' tail')) +
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr><w:ins w:id="4"${WHO('Bob')}/></w:trPr><w:tc>${p(ins(5, t('cell'), 'Bob'), markIns(6, 'Bob'))}</w:tc></w:tr></w:tbl>` +
      p('end'),
  );
  const list = revisionsOf(o.doc);
  assert.deepEqual(
    list.map((revision) => [revision.kind, revision.author, revision.text, revision.marks, revision.rows]),
    [
      ['insert', 'Ann', 'new¶more', 1, 0],
      ['insert', 'Bob', 'cell¶', 1, 1],
    ],
  );
  assert.deepEqual([list[0].start, list[0].end], [pos(0, 5), pos(1, 4)]);
});

test('revisions: the list is bounded and says when there is more; the text of a change is cut short', async () => {
  let content = '';
  for (let i = 0; i < 300; i++) content += ins(i + 1, t('x' + i), 'A' + i) + t('.');
  const o = await open(p(content) + p(ins(1000, t('y'.repeat(500)))));
  const some = listRevisions(o.doc, { limit: 10 });
  assert.equal(some.list.length, 10);
  assert.equal(some.more, true);
  const all = listRevisions(o.doc);
  assert.equal(all.list.length, 301);
  assert.equal(all.more, false);
  const long = all.list[300];
  assert.ok(long.text.length <= 81, String(long.text.length));
  assert.ok(long.text.endsWith('…'));
});

// ------------------------------------------------------- insertions, deletions

test('revisions: accepting an insertion leaves its text as plain text; rejecting takes the text out', async () => {
  const xml = p(t('one ') + ins(1, t('two ')) + t('three'));
  let o = await open(xml);
  const edit = acceptRevision(o.doc, first(o.doc));
  assert.equal(await applied(o, edit), `<w:p>${t('one two three')}</w:p>`);
  assert.deepEqual(revisionsOf(o.doc), []);
  // The caret is left at the end of what was accepted.
  assert.deepEqual(edit.selection.focus, pos(0, 8));

  o = await open(xml);
  const back = rejectRevision(o.doc, first(o.doc));
  assert.equal(await applied(o, back), `<w:p>${t('one three')}</w:p>`);
  assert.deepEqual(back.selection.focus, pos(0, 4));
});

test('revisions: accepting a deletion takes its text out; rejecting makes it text again', async () => {
  const xml = p(t('one ') + del(1, dt('two ')) + t('three'));
  let o = await open(xml);
  assert.equal(await applied(o, acceptRevision(o.doc, first(o.doc))), `<w:p>${t('one three')}</w:p>`);

  o = await open(xml);
  assert.equal(await applied(o, rejectRevision(o.doc, first(o.doc))), `<w:p>${t('one two three')}</w:p>`);
  assert.deepEqual(texts(o.doc), ['one two three']);
});

test('revisions: only the change that is named changes; the paragraphs beside it are copied from the file', async () => {
  const o = await open(p(ins(1, t('a'), 'Ann')) + p(t('x') + ins(2, t('b'), 'Bob') + del(3, dt('c'), 'Cid') + t('y'), '<w:jc w:val="center"/>') + p(del(4, dt('d'), 'Dan')));
  const before = o.doc.body.blocks.slice();
  run(o.history, rejectRevision(o.doc, first(o.doc, 'delete')));
  assert.ok(o.doc.body.blocks[0] === before[0] && o.doc.body.blocks[2] === before[2], 'the other paragraphs are the objects they were');
  assert.equal(await body(o), p(ins(1, t('a'), 'Ann')) + `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${t('x')}${ins(2, t('b'), 'Bob')}${t('cy')}</w:p>` + p(del(4, dt('d'), 'Dan')));
  assert.deepEqual(kinds(o.doc), ['insert', 'insert', 'delete']);
});

test('revisions: a rejected deletion of a field code is a field code again', async () => {
  const field =
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    del(1, '<w:r><w:delInstrText xml:space="preserve"> PAGE </w:delInstrText></w:r>') +
    ins(2, '<w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>') +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    t('1') +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const o = await open(p(field));
  /** The code the field at the start of the first paragraph is known by: what decides whether the engine computes it. */
  const code = (): string | undefined => {
    const [begin] = (o.doc.body.blocks[0] as Paragraph).children;
    const item = begin.type === 'run' ? begin.content[0] : undefined;
    return item && item.type === 'fieldChar' && item.field ? item.field.instruction : undefined;
  };
  const read = o.doc.body.blocks[0];
  run(o.history, rejectRevision(o.doc, first(o.doc, 'delete')));
  run(o.history, rejectRevision(o.doc, first(o.doc, 'insert')));
  // The field is the one it was before the two changes, for the layout too.
  assert.equal(code(), ' PAGE ');
  o.history.undo();
  o.history.undo();
  assert.ok(o.doc.body.blocks[0] === read);
  run(o.history, acceptAll(o.doc));
  assert.equal(code(), ' NUMPAGES ');
  o.history.undo();
  run(o.history, rejectAll(o.doc));
  const out = await body(o);
  assert.ok(out.indexOf('<w:instrText xml:space="preserve"> PAGE </w:instrText>') > 0, out);
  assert.ok(!REVISION_TAGS.test(out) && out.indexOf('NUMPAGES') < 0, out);
});

test('revisions: unknown attributes and elements of what stays are kept', async () => {
  const o = await open(
    p(`<w:ins w:id="1"${WHO('Ann')}><w:r w:rsidR="00AB12CD"><w:rPr><w:b/><w14:ligatures w14:val="all"/></w:rPr><w:t>new</w:t><w:unknownThing w:val="1"/></w:r></w:ins>`),
  );
  run(o.history, acceptRevision(o.doc, first(o.doc)));
  assert.equal(await body(o), '<w:p><w:r w:rsidR="00AB12CD"><w:rPr><w:b/><w14:ligatures w14:val="all"/></w:rPr><w:t>new</w:t><w:unknownThing w:val="1"/></w:r></w:p>');
});

// ------------------------------------------------------------ paragraph marks

test('revisions: accepting an inserted paragraph mark leaves two paragraphs; rejecting joins them', async () => {
  const xml = p(t('first'), '<w:pStyle w:val="Heading1"/>' + markIns(1)) + p(t('second'), '<w:jc w:val="center"/>');
  let o = await open(xml);
  assert.equal(first(o.doc).text, '¶');
  assert.equal(await applied(o, acceptRevision(o.doc, first(o.doc))), `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${t('first')}</w:p><w:p><w:pPr><w:jc w:val="center"/></w:pPr>${t('second')}</w:p>`);

  o = await open(xml);
  const edit = rejectRevision(o.doc, first(o.doc));
  // Word's rule: the mark that stays is the second paragraph's, and the joined paragraph is formatted like it.
  assert.equal(await applied(o, edit), `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${t('firstsecond')}</w:p>`);
  assert.deepEqual(edit.selection.focus, pos(0, 5));
});

test('revisions: accepting a deleted paragraph mark joins the paragraphs; rejecting keeps them apart', async () => {
  const xml = p(t('first'), markDel(1)) + '<w:bookmarkStart w:id="9" w:name="between"/><w:bookmarkEnd w:id="9"/>' + p(t('second'), markDel(2)) + p(t('third'));
  let o = await open(xml);
  assert.equal(revisionsOf(o.doc).length, 2);
  run(o.history, acceptAll(o.doc));
  assert.deepEqual(texts(o.doc), ['firstsecondthird']);
  // What stood between the paragraphs follows the joined one.
  assert.equal(await body(o), `<w:p>${t('firstsecondthird')}</w:p><w:bookmarkStart w:id="9" w:name="between"/><w:bookmarkEnd w:id="9"/>`);

  o = await open(xml);
  assert.equal(await applied(o, rejectAll(o.doc)), `<w:p>${t('first')}</w:p><w:bookmarkStart w:id="9" w:name="between"/><w:bookmarkEnd w:id="9"/><w:p>${t('second')}</w:p><w:p>${t('third')}</w:p>`);
});

test('revisions: an empty first paragraph takes the formatting of the one it joins', async () => {
  const o = await open(p(ins(1, t('typed')), '<w:jc w:val="right"/>' + markIns(2)) + p(t('old'), '<w:pStyle w:val="Quote"/>'));
  run(o.history, rejectAll(o.doc));
  assert.equal(await body(o), `<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr>${t('old')}</w:p>`);
});

test('revisions: a rejected mark with no paragraph to join goes with its empty paragraph, or stays as a plain mark', async () => {
  // Before a table: the empty paragraph that was put in goes.
  let o = await open(p('top') + p('', markIns(1)) + tbl(tr(tc('cell')), [2000]) + p('end'));
  run(o.history, rejectAll(o.doc));
  assert.deepEqual(texts(o.doc), ['top', 'cell', 'end']);
  assert.equal(o.doc.body.blocks.length, 3);

  // The last paragraph of a cell has to stay: its mark is a plain mark afterwards.
  o = await open(tbl(tr(tc(p(ins(1, t('typed')), markIns(2)))), [2000]) + p('end'));
  const result = rejectAll(o.doc);
  run(o.history, result);
  assert.deepEqual(texts(o.doc), ['', 'end']);
  assert.deepEqual(revisionsOf(o.doc), []);
  assert.ok(!REVISION_TAGS.test(await body(o)));
});

// ----------------------------------------------------------------------- moves

const MOVE =
  p(t('a ') + `<w:moveFromRangeStart w:id="1" w:name="move1"${WHO('Ann')}/><w:moveFrom w:id="2"${WHO('Ann')}>${t('moved ')}</w:moveFrom><w:moveFromRangeEnd w:id="1"/>` + t('b')) +
  p(t('c ') + `<w:moveToRangeStart w:id="3" w:name="move1"${WHO('Ann')}/><w:moveTo w:id="4"${WHO('Ann')}>${t('moved ')}</w:moveTo><w:moveToRangeEnd w:id="3"/>` + t('d'));

test('revisions: the two ends of a move are resolved together, and its range marks go', async () => {
  let o = await open(MOVE);
  assert.deepEqual(kinds(o.doc), ['moveFrom', 'moveTo']);
  assert.equal(await applied(o, acceptRevision(o.doc, first(o.doc, 'moveTo'))), `<w:p>${t('a b')}</w:p><w:p>${t('c moved d')}</w:p>`);
  assert.deepEqual(revisionsOf(o.doc), []);

  o = await open(MOVE);
  assert.equal(await applied(o, rejectRevision(o.doc, first(o.doc, 'moveFrom'))), `<w:p>${t('a moved b')}</w:p><w:p>${t('c d')}</w:p>`);
});

test('revisions: an end of a move without its partner is resolved alone, like an insertion or a deletion', async () => {
  let o = await open(p(t('a ') + `<w:moveFrom w:id="2"${WHO('Ann')}>${t('gone ')}</w:moveFrom>` + t('b')) + p(`<w:moveTo w:id="4"${WHO('Ann')}>${t('came')}</w:moveTo>`));
  run(o.history, acceptRevision(o.doc, first(o.doc, 'moveFrom')));
  assert.deepEqual(texts(o.doc), ['a b', 'came']);
  assert.deepEqual(kinds(o.doc), ['moveTo']);
  run(o.history, rejectRevision(o.doc, first(o.doc, 'moveTo')));
  assert.deepEqual(texts(o.doc), ['a b', '']);
});

// ------------------------------------------------------------------ formatting

test('revisions: accepting a formatting change keeps the formatting; rejecting brings the old one back', async () => {
  const change = `<w:rPrChange w:id="1"${WHO('Ann')}><w:rPr><w:i/><w:color w:val="FF0000"/></w:rPr></w:rPrChange>`;
  const xml = p(`<w:r><w:rPr><w:b/>${change}</w:rPr><w:t>text</w:t></w:r>`);
  let o = await open(xml);
  const [revision] = revisionsOf(o.doc);
  assert.deepEqual([revision.kind, revision.author, revision.text], ['format', 'Ann', 'text']);
  assert.deepEqual([revision.start, revision.end], [pos(0, 0), pos(0, 4)]);
  assert.equal(await applied(o, acceptRevision(o.doc, revision)), '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>text</w:t></w:r></w:p>');

  o = await open(xml);
  assert.equal(await applied(o, rejectRevision(o.doc, first(o.doc))), '<w:p><w:r><w:rPr><w:i/><w:color w:val="FF0000"/></w:rPr><w:t>text</w:t></w:r></w:p>');

  // No old formatting at all: the run had none.
  o = await open(p(`<w:r><w:rPr><w:b/><w:rPrChange w:id="1"${WHO('Ann')}><w:rPr/></w:rPrChange></w:rPr><w:t>text</w:t></w:r>`));
  assert.equal(await applied(o, rejectRevision(o.doc, first(o.doc))), `<w:p>${t('text')}</w:p>`);
});

test('revisions: a change of paragraph formatting, and of the formatting of a mark', async () => {
  const pChange = `<w:pPrChange w:id="1"${WHO('Ann')}><w:pPr><w:pStyle w:val="Quote"/><w:ind w:left="720"/></w:pPr></w:pPrChange>`;
  const mChange = `<w:rPr><w:b/><w:rPrChange w:id="2"${WHO('Bob')}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr>`;
  const xml = p(t('text'), `<w:jc w:val="center"/>${mChange}${pChange}`);
  let o = await open(xml);
  assert.deepEqual(kinds(o.doc), ['paragraphFormat', 'format']);
  run(o.history, acceptAll(o.doc));
  assert.equal(await body(o), `<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:b/></w:rPr></w:pPr>${t('text')}</w:p>`);

  o = await open(xml);
  assert.equal(await applied(o, rejectRevision(o.doc, first(o.doc, 'paragraphFormat'))), `<w:p><w:pPr><w:pStyle w:val="Quote"/><w:ind w:left="720"/>${mChange}</w:pPr>${t('text')}</w:p>`);
  run(o.history, rejectRevision(o.doc, first(o.doc, 'format')));
  assert.equal(await body(o), `<w:p><w:pPr><w:pStyle w:val="Quote"/><w:ind w:left="720"/><w:rPr><w:i/></w:rPr></w:pPr>${t('text')}</w:p>`);
});

test('revisions: a change of section formatting, in a paragraph and at the end of the body', async () => {
  const old = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';
  const inner = `<w:sectPr><w:pgSz w:w="9000" w:h="9000"/><w:sectPrChange w:id="1"${WHO('Ann')}>${old}</w:sectPrChange></w:sectPr>`;
  const last = `<w:sectPr><w:pgSz w:w="8000" w:h="8000"/><w:sectPrChange w:id="2"${WHO('Ann')}>${old}</w:sectPrChange></w:sectPr>`;
  const make = async (): Promise<Opened> => {
    const bytes = await buildSpec({ body: p(t('one'), inner) + p('two') + last, styles: STYLES });
    const pkg = await OpcPackage.open(bytes);
    const doc = await readDocx(pkg);
    return { pkg, doc, history: new DocHistory<DocRange | undefined>(doc), bytes };
  };
  let o = await make();
  assert.deepEqual(kinds(o.doc), ['sectionFormat', 'sectionFormat']);
  run(o.history, acceptAll(o.doc));
  let out = await written(o);
  assert.ok(out.indexOf('sectPrChange') < 0 && out.indexOf('w:w="9000"') > 0 && out.indexOf('w:w="8000"') > 0 && out.indexOf('12240') < 0, out);
  o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);

  o = await make();
  run(o.history, rejectAll(o.doc));
  out = await written(o);
  assert.ok(out.indexOf('sectPrChange') < 0 && out.indexOf('9000') < 0 && out.indexOf('8000') < 0, out);
  assert.equal(out.split('<w:pgSz w:w="12240" w:h="15840"/>').length, 3, out);
  assert.deepEqual(revisionsOf(o.doc), []);
});

test('revisions: changes of the formatting of a table, a row and a cell', async () => {
  const table =
    `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:jc w:val="center"/><w:tblPrChange w:id="1"${WHO('Ann')}><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr></w:tblPrChange></w:tblPr>` +
    '<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
    `<w:tr><w:trPr><w:cantSplit/><w:trPrChange w:id="2"${WHO('Ann')}><w:trPr><w:tblHeader/></w:trPr></w:trPrChange></w:trPr>` +
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:vAlign w:val="center"/><w:tcPrChange w:id="3"${WHO('Ann')}><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr></w:tcPrChange></w:tcPr>${p('cell')}</w:tc></w:tr></w:tbl>`;
  let o = await open(table + p('end'));
  // One author, one date, side by side: one change, as Word steps over it.
  assert.deepEqual(kinds(o.doc), ['tableFormat']);
  run(o.history, acceptAll(o.doc));
  let out = await body(o);
  assert.ok(!REVISION_TAGS.test(out) && out.indexOf('<w:jc w:val="center"/>') > 0 && out.indexOf('<w:cantSplit/>') > 0 && out.indexOf('<w:vAlign w:val="center"/>') > 0, out);

  o = await open(table + p('end'));
  run(o.history, rejectAll(o.doc));
  out = await body(o);
  assert.ok(!REVISION_TAGS.test(out), out);
  assert.ok(out.indexOf('<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>') > 0 && out.indexOf('<w:trPr><w:tblHeader/></w:trPr>') > 0 && out.indexOf('<w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr>') > 0, out);
  o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);
});

// ----------------------------------------------------------------------- rows

const rowOf = (text: string, trPr = '') => `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${tc(text)}</w:tr>`;
const ROWS = tbl(rowOf('one') + rowOf('two', `<w:ins w:id="1"${WHO('Ann')}/>`) + rowOf('three', `<w:del w:id="2"${WHO('Bob')}/>`), [2000]);

test('revisions: an inserted row stays when accepted and goes when rejected; a deleted row the other way round', async () => {
  let o = await open(p('top') + ROWS + p('end'));
  assert.deepEqual(
    revisionsOf(o.doc).map((revision) => [revision.kind, revision.rows, revision.author]),
    [
      ['insert', 1, 'Ann'],
      ['delete', 1, 'Bob'],
    ],
  );
  run(o.history, acceptAll(o.doc));
  assert.deepEqual(texts(o.doc), ['top', 'one', 'two', 'end']);
  assert.ok(!REVISION_TAGS.test(await body(o)));

  o = await open(p('top') + ROWS + p('end'));
  const out = await applied(o, rejectAll(o.doc));
  assert.deepEqual(texts(o.doc), ['top', 'one', 'three', 'end']);
  assert.ok(!REVISION_TAGS.test(out));
});

test('revisions: the table goes with its last row, and a paragraph stays where one is needed', async () => {
  const table = tbl(rowOf('new', `<w:ins w:id="1"${WHO('Ann')}/>`) + rowOf('newer', `<w:ins w:id="2"${WHO('Ann')}/>`), [2000]);
  let o = await open(p('top') + table + p('end'));
  // The text of the cells is no insertion, so each row is a change of its own.
  assert.equal(revisionsOf(o.doc).length, 2);
  run(o.history, rejectRevision(o.doc, first(o.doc)));
  assert.deepEqual(texts(o.doc), ['top', 'newer', 'end']);
  run(o.history, rejectRevision(o.doc, first(o.doc)));
  assert.deepEqual(texts(o.doc), ['top', 'end']);
  assert.equal(o.doc.body.blocks.length, 2);

  // Between two tables an empty paragraph takes its place.
  o = await open(tbl(rowOf('a'), [2000]) + table + tbl(rowOf('b'), [2000]) + p('end'));
  run(o.history, rejectAll(o.doc));
  assert.deepEqual(texts(o.doc), ['a', '', 'b', 'end']);
  assert.deepEqual(
    o.doc.body.blocks.map((block) => block.type),
    ['table', 'paragraph', 'table', 'paragraph'],
  );
});

test('revisions: a vertically merged cell that began in a row that goes begins in the next one', async () => {
  const cell = (text: string, merge: string) => `<w:tc><w:tcPr><w:vMerge${merge}/></w:tcPr>${p(text)}</w:tc>`;
  const o = await open(tbl(`<w:tr><w:trPr><w:del w:id="1"${WHO('Ann')}/></w:trPr>${cell('a', ' w:val="restart"')}</w:tr><w:tr>${cell('', '')}</w:tr>`, [2000]) + p('end'));
  run(o.history, acceptAll(o.doc));
  const table = o.doc.body.blocks[0] as Table;
  assert.equal(table.rows.length, 1);
  assert.equal(table.rows[0].cells[0].props.vMerge, 'restart');
});

// ------------------------------------------------- what the model does not hold

test('revisions: changes the model only keeps as XML are listed as "other", cannot be resolved one by one, and are skipped and counted by "all"', async () => {
  // (Inserted and deleted cells are in the model and can be resolved: test/doc-track-tables.test.ts.)
  const cellIns = `<w:tc><w:tcPr><w:cellMerge w:id="1"${WHO('Ann')} w:vMerge="rest"/></w:tcPr>${p('cell')}</w:tc>`;
  const o = await open(p(ins(2, t('typed'), 'Bob')) + tbl(`<w:tr>${cellIns}</w:tr>`, [2000]) + p('end'));
  const list = revisionsOf(o.doc);
  assert.deepEqual(
    list.map((revision) => [revision.kind, revision.canAccept, revision.canReject]),
    [
      ['insert', true, true],
      ['other', false, false],
    ],
  );
  refused(() => acceptRevision(o.doc, list[1]), REVISION_UNSUPPORTED);
  refused(() => rejectRevision(o.doc, list[1]), REVISION_UNSUPPORTED);
  const all = acceptAll(o.doc);
  assert.deepEqual([all.resolved, all.skipped], [1, 1]);
  run(o.history, all);
  assert.deepEqual(kinds(o.doc), ['other']);
  assert.ok((await body(o)).indexOf('<w:cellMerge w:id="1"') > 0);
  // Nothing left that can be resolved: nothing to do, and it says so.
  const none = rejectAll(o.doc);
  assert.deepEqual([none.ops.length, none.resolved, none.skipped], [0, 0, 1]);
});

test('revisions: a rejected formatting change of a cell that also carries a change kept as XML is refused; accepting it is not', async () => {
  const cell = `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:cellMerge w:id="1"${WHO('Ann')} w:vMerge="rest"/><w:tcPrChange w:id="2"${WHO('Bob')}><w:tcPr><w:tcW w:w="900" w:type="dxa"/></w:tcPr></w:tcPrChange></w:tcPr>${p('cell')}</w:tc>`;
  const o = await open(tbl(`<w:tr>${cell}</w:tr>`, [2000]) + p('end'));
  const format = first(o.doc, 'tableFormat');
  assert.deepEqual([format.canAccept, format.canReject], [true, false]);
  refused(() => rejectRevision(o.doc, format), REVISION_UNSUPPORTED);
  const all = rejectAll(o.doc);
  assert.deepEqual([all.resolved, all.skipped], [0, 2]);
  run(o.history, acceptRevision(o.doc, format));
  const out = await body(o);
  assert.ok(out.indexOf('tcPrChange') < 0 && out.indexOf('<w:cellMerge w:id="1"') > 0 && out.indexOf('w:w="2000"') > 0, out);
});

test('revisions: a change that is gone from the document is refused', async () => {
  const o = await open(p(ins(1, t('a'))) + p(del(2, dt('b'), 'Bob')));
  const [a, b] = revisionsOf(o.doc);
  run(o.history, acceptRevision(o.doc, a));
  refused(() => acceptRevision(o.doc, a), REVISION_GONE);
  // A change the edit did not touch is still the change it was.
  run(o.history, acceptRevision(o.doc, b));
  assert.deepEqual(texts(o.doc), ['a', '']);
});

// ---------------------------------------------------------------------- nesting

test('revisions: a deletion inside an insertion, resolved in either order', async () => {
  const xml = p(t('a ') + `<w:ins w:id="1"${WHO('Ann')}>${t('new ')}<w:del w:id="2"${WHO('Bob')}>${dt('cut ')}</w:del>${t('text ')}</w:ins>` + t('b'));
  const outcome = async (steps: [kind: string, accept: boolean][]): Promise<string> => {
    const o = await open(xml);
    assert.deepEqual(
      revisionsOf(o.doc).map((revision) => [revision.kind, revision.text]),
      [
        ['insert', 'new text '],
        ['delete', 'cut '],
      ],
    );
    for (const [kind, accept] of steps) {
      const found = revisionsOf(o.doc).find((revision) => revision.kind === kind);
      if (!found) continue;
      run(o.history, accept ? acceptRevision(o.doc, found) : rejectRevision(o.doc, found));
      checkInvariants(o.doc);
    }
    assert.deepEqual(revisionsOf(o.doc), []);
    return (await body(o)) + '|' + texts(o.doc).join('/');
  };
  // Rejecting the insertion takes everything in it, whatever was decided about the deletion inside.
  for (const inner of [true, false]) {
    assert.equal(await outcome([['insert', false], ['delete', inner]]), `<w:p>${t('a b')}</w:p>|a b`);
    assert.equal(await outcome([['delete', inner], ['insert', false]]), `<w:p>${t('a b')}</w:p>|a b`);
  }
  assert.equal(await outcome([['insert', true], ['delete', true]]), await outcome([['delete', true], ['insert', true]]));
  assert.equal(await outcome([['insert', true], ['delete', true]]), `<w:p>${t('a new text b')}</w:p>|a new text b`);
  assert.equal(await outcome([['insert', true], ['delete', false]]), await outcome([['delete', false], ['insert', true]]));
  assert.equal(await outcome([['delete', false], ['insert', true]]), `<w:p>${t('a new cut text b')}</w:p>|a new cut text b`);
});

// ------------------------------------------- notes, links, controls and fields

const NOTES = (inner: string) =>
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' + inner;
const NOTE = (id: number, text: string) => `<w:footnote w:id="${id}"><w:p><w:r><w:footnoteRef/></w:r>${t(text)}</w:p></w:footnote>`;
const REF = (id: number) => `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;

test('revisions: a note goes with its reference: when an inserted reference is rejected, when a deleted one is accepted', async () => {
  const xml = p(t('a') + ins(1, REF(2)) + del(2, REF(3), 'Bob') + t('b'));
  let o = await open(xml, { footnotes: NOTES(NOTE(2, 'two') + NOTE(3, 'three')) });
  run(o.history, rejectRevision(o.doc, first(o.doc, 'insert')));
  assert.deepEqual(Array.from(o.doc.footnotes!.items.keys()), [-1, 0, 3]);
  run(o.history, acceptRevision(o.doc, first(o.doc, 'delete')));
  assert.deepEqual(Array.from(o.doc.footnotes!.items.keys()), [-1, 0]);
  assert.equal(await body(o), `<w:p>${t('ab')}</w:p>`);
  const notes = await written(o, '/word/footnotes.xml');
  assert.ok(notes.indexOf('two') < 0 && notes.indexOf('three') < 0, notes);
  // Undone, both notes are back and the file is what it was.
  o.history.undo();
  o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);

  // The other way round both notes stay.
  o = await open(xml, { footnotes: NOTES(NOTE(2, 'two') + NOTE(3, 'three')) });
  run(o.history, acceptRevision(o.doc, first(o.doc, 'insert')));
  run(o.history, rejectRevision(o.doc, first(o.doc, 'delete')));
  assert.deepEqual(Array.from(o.doc.footnotes!.items.keys()), [-1, 0, 2, 3]);
  assert.equal(await body(o), `<w:p>${t('a')}${REF(2)}${REF(3)}${t('b')}</w:p>`);
});

test('revisions: changes inside notes and headers are listed after the body and resolved like any other', async () => {
  const o = await open(p(t('body') + REF(2)) + p(ins(9, t('x'))), {
    footnotes: NOTES(`<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r>${t('note ')}${ins(5, t('added'), 'Bob')}</w:p></w:footnote>`),
    headers: { rId7: p(t('head ') + del(6, dt('gone'), 'Cid')) },
    body: p(t('body') + REF(2)) + p(ins(9, t('x'))) + `<w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:pgSz w:w="9000" w:h="9000"/></w:sectPr>`,
  });
  assert.deepEqual(
    revisionsOf(o.doc).map((revision) => [revision.story, revision.kind, revision.text]),
    [
      ['body', 'insert', 'x'],
      ['header:rId7', 'delete', 'gone'],
      ['footnote:2', 'insert', 'added'],
    ],
  );
  run(o.history, rejectRevision(o.doc, revisionsOf(o.doc)[2]));
  run(o.history, rejectRevision(o.doc, revisionsOf(o.doc)[1]));
  assert.deepEqual(texts(o.doc, 'footnote:2'), ['note ']);
  assert.deepEqual(texts(o.doc, 'header:rId7'), ['head gone']);
  assert.ok((await written(o, '/word/header1.xml')).indexOf(t('head gone')) > 0);
  assert.ok(!REVISION_TAGS.test(await written(o, '/word/footnotes.xml')));
  o.history.undo();
  o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);
});

test('revisions: a link that loses all its text goes, a content control and a simple field stay, empty', async () => {
  const link = `<w:hyperlink w:anchor="here">${ins(1, t('link text'))}<w:bookmarkStart w:id="7" w:name="kept"/><w:bookmarkEnd w:id="7"/></w:hyperlink>`;
  const control = `<w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent>${ins(2, t('in control'))}</w:sdtContent></w:sdt>`;
  const field = `<w:fldSimple w:instr=" AUTHOR ">${ins(3, t('result'))}</w:fldSimple>`;
  const o = await open(p(t('a') + link + control + field + t('b')));
  run(o.history, rejectAll(o.doc));
  assert.equal(
    await body(o),
    `<w:p>${t('a')}<w:bookmarkStart w:id="7" w:name="kept"/><w:bookmarkEnd w:id="7"/><w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent/></w:sdt><w:fldSimple w:instr=" AUTHOR "></w:fldSimple>${t('b')}</w:p>`,
  );

  // Accepted, each of them holds plain text.
  const again = await open(p(t('a') + link + control + field + t('b')));
  run(again.history, acceptAll(again.doc));
  assert.equal(
    await body(again),
    `<w:p>${t('a')}<w:hyperlink w:anchor="here">${t('link text')}<w:bookmarkStart w:id="7" w:name="kept"/><w:bookmarkEnd w:id="7"/></w:hyperlink><w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent>${t('in control')}</w:sdtContent></w:sdt><w:fldSimple w:instr=" AUTHOR ">${t('result')}</w:fldSimple>${t('b')}</w:p>`,
  );
});

test('revisions: an inserted field goes as a whole; a change that would leave half a field is refused, and skipped by "all"', async () => {
  const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>';
  const end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  let o = await open(p(t('a') + ins(1, begin + t('today') + end) + t('b')));
  assert.equal(await applied(o, rejectRevision(o.doc, first(o.doc))), `<w:p>${t('ab')}</w:p>`);

  o = await open(p(t('a') + ins(1, begin) + t('today') + end + t('b')));
  refused(() => rejectRevision(o.doc, first(o.doc)), REVISION_UNSUPPORTED);
  const all = rejectAll(o.doc);
  assert.deepEqual([all.ops.length, all.resolved, all.skipped], [0, 0, 1]);
  // Accepting it breaks nothing.
  run(o.history, acceptRevision(o.doc, first(o.doc)));
  assert.equal(await body(o), `<w:p>${t('a')}${begin}${t('today')}${end}${t('b')}</w:p>`);
});

test('revisions: a picture and a comment reference in a rejected insertion: the picture goes, the comment marks stay', async () => {
  const comment = '<w:commentRangeStart w:id="4"/>' + t('said') + '<w:commentRangeEnd w:id="4"/><w:r><w:commentReference w:id="4"/></w:r>';
  const o = await open(p(t('a') + ins(1, comment + '<w:r><w:pict><v:rect style="width:10pt;height:10pt"/></w:pict></w:r>') + t('b')), {
    comments: `<w:comment w:id="4" w:author="Ann"><w:p>${t('note')}</w:p></w:comment>`,
  });
  run(o.history, rejectAll(o.doc));
  assert.equal(await body(o), `<w:p>${t('a')}<w:commentRangeStart w:id="4"/><w:commentRangeEnd w:id="4"/><w:r><w:commentReference w:id="4"/></w:r>${t('b')}</w:p>`);
});

// ------------------------------------------------------- protection, tracking

test('revisions: a document whose protection is enforced lets no change be accepted or rejected', async () => {
  for (const edit of ['trackedChanges', 'readOnly', 'comments', 'forms']) {
    const o = await open(p(ins(1, t('a'))), { settings: `<w:documentProtection w:edit="${edit}" w:enforcement="1"/>` });
    const [revision] = revisionsOf(o.doc);
    assert.deepEqual([revision.canAccept, revision.canReject], [false, false], edit);
    refused(() => acceptRevision(o.doc, revision), REVISIONS_PROTECTED);
    refused(() => rejectRevision(o.doc, revision), REVISIONS_PROTECTED);
    refused(() => acceptAll(o.doc), REVISIONS_PROTECTED);
    refused(() => rejectAll(o.doc), REVISIONS_PROTECTED);
  }
  // A protection that is not enforced restricts nothing.
  const free = await open(p(ins(1, t('a'))), { settings: '<w:documentProtection w:edit="trackedChanges" w:enforcement="0"/>' });
  run(free.history, acceptAll(free.doc));
  assert.deepEqual(revisionsOf(free.doc), []);
});

test('revisions: accepting and rejecting are not tracked themselves, also where every edit is', async () => {
  const o = await open(p(t('one ') + ins(1, t('two ')) + del(2, dt('three '), 'Bob') + t('four')), { settings: '<w:trackRevisions/>' });
  const me = { author: 'Me', date: '2026-10-08T10:00:00Z' };
  run(o.history, insertText(o.doc, pos(0, 0), 'X', undefined, me));
  run(o.history, deleteRange(o.doc, range(pos(0, 9), pos(0, 13)), me));
  assert.deepEqual(
    revisionsOf(o.doc).map((revision) => `${revision.kind}:${revision.author}:${revision.text}`),
    ['insert:Me:X', 'insert:Ann:two ', 'delete:Bob:three ', 'delete:Me:four'],
  );
  run(o.history, rejectRevision(o.doc, first(o.doc, 'insert')));
  run(o.history, acceptRevision(o.doc, revisionsOf(o.doc)[2]));
  assert.deepEqual(
    revisionsOf(o.doc).map((revision) => `${revision.kind}:${revision.author}:${revision.text}`),
    ['insert:Ann:two ', 'delete:Bob:three '],
  );
  run(o.history, acceptAll(o.doc));
  assert.equal(await body(o), `<w:p>${t('one two ')}</w:p>`);
});

// ------------------------------------------------------------------------- all

test('revisions: "accept all" leaves the text with the insertions and without the deletions, "reject all" the other way round, and no mark of a change', async () => {
  const header = p(t('h') + ins(20, t('+')) + del(21, dt('-')));
  const note = `<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r>${t('n')}${ins(22, t('+'))}${del(23, dt('-'))}</w:p></w:footnote>`;
  const content =
    p(t('a') + ins(1, t('+')) + del(2, dt('-')) + REF(2), markIns(3)) +
    p(t('b') + `<w:moveFrom w:id="4"${WHO('Ann')}>${t('m')}</w:moveFrom>`, markDel(5)) +
    p(t('c') + `<w:moveTo w:id="6"${WHO('Ann')}>${t('m')}</w:moveTo>` + `<w:r><w:rPr><w:b/><w:rPrChange w:id="7"${WHO('Ann')}><w:rPr/></w:rPrChange></w:rPr><w:t>f</w:t></w:r>`) +
    tbl(rowOf('r1') + rowOf('r2', `<w:ins w:id="8"${WHO('Ann')}/>`) + rowOf('r3', `<w:del w:id="9"${WHO('Ann')}/>`), [2000]) +
    p('end');
  const spec = {
    body: content + '<w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:pgSz w:w="9000" w:h="9000"/></w:sectPr>',
    footnotes: NOTES(note),
    headers: { rId7: header },
  };
  const everything = (doc: DocModel): string => [texts(doc).join('/'), texts(doc, 'header:rId7').join('/'), texts(doc, 'footnote:2').join('/')].join(' | ');

  let o = await open('', spec);
  const accepted = acceptAll(o.doc);
  assert.equal(accepted.skipped, 0);
  run(o.history, accepted);
  checkInvariants(o.doc);
  assert.equal(everything(o.doc), 'a+/bcmf/r1/r2/end | h+ | n+');
  for (const part of ['/word/document.xml', '/word/header1.xml', '/word/footnotes.xml']) assert.ok(!REVISION_TAGS.test(await written(o, part)), part);
  assert.deepEqual(revisionsOf(o.doc), []);
  o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);

  o = await open('', spec);
  run(o.history, rejectAll(o.doc));
  checkInvariants(o.doc);
  assert.equal(everything(o.doc), 'a-bm/cf/r1/r3/end | h- | n-');
  for (const part of ['/word/document.xml', '/word/header1.xml', '/word/footnotes.xml']) assert.ok(!REVISION_TAGS.test(await written(o, part)), part);
  // The formatting that was changed is the old one again.
  assert.ok((await body(o)).indexOf('<w:b/>') < 0);
});

test('revisions: "all" takes a filter: the changes of one author only', async () => {
  const o = await open(p(ins(1, t('a'), 'Ann') + t(' ') + ins(2, t('b'), 'Bob') + t(' ') + del(3, dt('c'), 'Ann')));
  const result = rejectAll(o.doc, (revision) => revision.author === 'Ann');
  assert.deepEqual([result.resolved, result.skipped], [2, 0]);
  run(o.history, result);
  assert.equal(await body(o), `<w:p>${t(' ')}${ins(2, t('b'), 'Bob')}${t(' c')}</w:p>`);
});

test('revisions: a document without changes has an empty list and "all" does nothing', async () => {
  const o = await open(p('plain') + tbl(tr(tc('cell')), [2000]) + p('end'));
  assert.deepEqual(listRevisions(o.doc), { list: [], more: false });
  const all = acceptAll(o.doc);
  assert.deepEqual([all.ops.length, all.resolved, all.skipped], [0, 0, 0]);
});

test('revisions: paragraphs of a cell: changes in a table are found by their path and resolved in place', async () => {
  const o = await open(p('top') + tbl(tr(tc(p(t('x') + ins(1, t('y'))) + p(del(2, dt('z'), 'Bob') + t('w'))) + tc('other')), [2000, 2000]) + p('end'));
  const list = revisionsOf(o.doc);
  assert.deepEqual(
    list.map((revision) => [revision.start.path, revision.start.offset]),
    [
      [[1, 0, 0, 0], 1],
      [[1, 0, 0, 1], 0],
    ],
  );
  run(o.history, rejectRevision(o.doc, list[0]));
  run(o.history, rejectRevision(o.doc, first(o.doc)));
  assert.deepEqual(texts(o.doc), ['top', 'x', 'zw', 'other', 'end']);
  const table = o.doc.body.blocks[1] as Table;
  assert.equal((table.rows[0].cells[0].blocks[0] as Paragraph).children.length, 1);
});

// ------------------------------------------------------------ converted files

test('revisions: the changes of an RTF and of a doc file, opened for editing as docx, are accepted and rejected like any other', async () => {
  const rtf = readRtf('{\\rtf1{\\*\\revtbl{Unknown;}{Ann;}{Bob;}}keep {\\deleted\\revauthdel1 gone }{\\revised\\revauth2 new }tail\\par}');
  const builder = new DocBuilder();
  const text = 'kept gone new end\r';
  builder.text(text);
  builder.papxAll(text);
  builder.table(FC.SttbfRMark, sttb(['Unknown', 'Ann', 'Bob']));
  builder.chpx(5, 10, sprm(0x0800, 1), sprm(0x4863, u16le(1)));
  builder.chpx(10, 14, sprm(0x0801, 1), sprm(0x4804, u16le(2)));
  const cases: [name: string, source: DocModel, accepted: string, rejected: string][] = [
    ['rtf', rtf, 'keep new tail', 'keep gone tail'],
    ['doc', builder.read(), 'kept new end', 'kept gone end'],
  ];
  for (const [name, source, accepted, rejected] of cases) {
    for (const accept of [true, false]) {
      // What mount({ convert: true }) does: the model is written as a docx, and that file is the one that is edited.
      const pkg = await OpcPackage.open(await convertToDocx(source));
      const doc = await readDocx(pkg);
      const history = new DocHistory<DocRange | undefined>(doc);
      assert.deepEqual(
        revisionsOf(doc).map((revision) => [revision.kind, revision.author]),
        [
          ['delete', 'Ann'],
          ['insert', 'Bob'],
        ],
        name,
      );
      // One by one, then the rest.
      run(history, accept ? acceptRevision(doc, first(doc)) : rejectRevision(doc, first(doc)));
      run(history, accept ? acceptAll(doc) : rejectAll(doc));
      assert.deepEqual(texts(doc), [accept ? accepted : rejected], name);
      await writeDocx(pkg, doc);
      const out = decoder.decode(await pkg.readPart('/word/document.xml'));
      parseXml(out);
      assert.ok(!REVISION_TAGS.test(out), out);
      const again = await readDocx(await OpcPackage.open(await pkg.save()));
      assert.deepEqual(texts(again), [accept ? accepted : rejected], name);
      assert.deepEqual(revisionsOf(again), []);
    }
  }
});

// ---------------------------------------------------- what stands between changes

test('revisions: a change of the formatting of a table alone is found and resolved', async () => {
  const table = `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:tblPrChange w:id="1"${WHO('Ann')}><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr></w:tblPrChange></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${tc('cell')}</w:tr></w:tbl>`;
  for (const accept of [true, false]) {
    const o = await open(p('top') + table + p('end'));
    assert.deepEqual(kinds(o.doc), ['tableFormat']);
    const out = await applied(o, accept ? acceptRevision(o.doc, first(o.doc)) : rejectRevision(o.doc, first(o.doc)));
    assert.ok(out.indexOf('tblPrChange') < 0, out);
    assert.ok(out.indexOf(accept ? '<w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr>' : '<w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>') > 0, out);
    assert.deepEqual(revisionsOf(o.doc), []);
  }
});

test('revisions: content that is written back whole between two changes is left alone by "all", and the changes inside it are counted', async () => {
  const alternate =
    '<mc:AlternateContent><mc:Choice Requires="wps">' +
    p(t('choice ') + ins(5, t('inside'), 'Eve')) +
    '</mc:Choice><mc:Fallback>' +
    p(t('fallback')) +
    '</mc:Fallback></mc:AlternateContent>';
  const o = await open(p(ins(1, t('first'))) + alternate + p(del(2, dt('second'), 'Bob') + t('x')) + p('end'));
  assert.deepEqual(
    revisionsOf(o.doc).map((revision) => [revision.kind, revision.count]),
    [
      ['insert', 1],
      ['other', 1],
      ['delete', 1],
    ],
  );
  const all = acceptAll(o.doc);
  assert.deepEqual([all.resolved, all.skipped], [2, 1]);
  const out = await applied(o, all);
  // The last paragraph is copied from the file as it was written there.
  assert.equal(out, p(t('first')) + alternate + p(t('x')) + p('end'));
  assert.deepEqual(kinds(o.doc), ['other']);
});

test('revisions: what stands between two paragraphs that are joined is not written again when it need not be', async () => {
  const o = await open(p(ins(1, t('a'))) + p('plain one') + p('plain two') + p(del(2, dt('b'), 'Bob') + t('c')));
  const before = o.doc.body.blocks.slice();
  const edit = acceptAll(o.doc);
  // Two paragraphs change: each is replaced by itself, the two between them are not touched.
  assert.deepEqual(
    edit.ops.map((op) => (op.type === 'replaceBlocks' ? [op.at, op.remove, op.insert.length] : op.type)),
    [
      [3, 1, 1],
      [0, 1, 1],
    ],
  );
  run(o.history, edit);
  assert.ok(o.doc.body.blocks[1] === before[1] && o.doc.body.blocks[2] === before[2]);
});

test('revisions: an insertion over several paragraphs is rejected as one; a mark that was inserted and then deleted takes two decisions', async () => {
  let o = await open(p(t('keep ') + ins(1, t('one')), markIns(2)) + p(ins(3, t('two')), markIns(4)) + p(ins(5, t('three')) + t(' tail')) + p('end'));
  assert.equal(revisionsOf(o.doc).length, 1);
  assert.equal(await applied(o, rejectRevision(o.doc, first(o.doc))), `<w:p>${t('keep  tail')}</w:p>` + p('end'));

  // Inserted by Ann, deleted by Bob: accepting the insertion leaves the deletion; accepting that joins the paragraphs.
  o = await open(p(t('a'), `<w:rPr><w:ins w:id="1"${WHO('Ann')}/><w:del w:id="2"${WHO('Bob')}/></w:rPr>`) + p('b'));
  assert.deepEqual(kinds(o.doc), ['insert', 'delete']);
  run(o.history, acceptRevision(o.doc, first(o.doc, 'insert')));
  assert.equal(await body(o), `<w:p><w:pPr><w:rPr><w:del w:id="2"${WHO('Bob')}/></w:rPr></w:pPr>${t('a')}</w:p>` + p('b'));
  run(o.history, acceptRevision(o.doc, first(o.doc)));
  assert.deepEqual(texts(o.doc), ['ab']);
});

test('revisions: changes in paragraphs of a content control around blocks, and in a header two sections share', async () => {
  const control = `<w:sdt><w:sdtPr><w:id w:val="77"/></w:sdtPr><w:sdtContent>${p(t('in ') + ins(1, t('control')))}${p(del(2, dt('gone'), 'Bob') + t('stays'))}</w:sdtContent></w:sdt>`;
  const sect = '<w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:pgSz w:w="9000" w:h="9000"/></w:sectPr>';
  const spec = {
    body: p(t('one'), '<w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:pgSz w:w="9000" w:h="9000"/></w:sectPr>') + control + p('end') + sect,
    headers: { rId7: p(t('head ') + ins(3, t('new'), 'Cid')) },
  };
  let o = await open('', spec);
  // The header is one story, however many sections name it.
  assert.deepEqual(
    revisionsOf(o.doc).map((revision) => [revision.story, revision.kind]),
    [
      ['body', 'insert'],
      ['body', 'delete'],
      ['header:rId7', 'insert'],
    ],
  );
  run(o.history, rejectRevision(o.doc, first(o.doc, 'delete')));
  const out = await body(o);
  // One control, with its id, around both paragraphs; the paragraph that was not touched keeps its change.
  assert.equal(out.split('<w:sdt>').length, 2, out);
  assert.ok(out.indexOf(`<w:sdt><w:sdtPr><w:id w:val="77"/></w:sdtPr><w:sdtContent>${p(t('in ') + ins(1, t('control')))}<w:p>${t('gonestays')}</w:p></w:sdtContent></w:sdt>`) > 0, out);
  run(o.history, acceptAll(o.doc));
  assert.deepEqual(revisionsOf(o.doc), []);
  assert.ok((await written(o, '/word/header1.xml')).indexOf(t('head new')) > 0);
  while (o.history.canUndo) o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);

  o = await open('', spec);
  run(o.history, rejectAll(o.doc));
  assert.deepEqual([texts(o.doc), texts(o.doc, 'header:rId7')], [['one', 'in ', 'gonestays', 'end'], ['head ']]);
});

test('revisions: range marks of a move that stand between blocks go with the move; an inserted row takes its note along', async () => {
  const moved =
    `<w:moveFromRangeStart w:id="1" w:name="m"${WHO('Ann')}/>` +
    p(`<w:moveFrom w:id="2"${WHO('Ann')}>${t('moved')}</w:moveFrom>`, `<w:rPr><w:moveFrom w:id="3"${WHO('Ann')}/></w:rPr>`) +
    '<w:moveFromRangeEnd w:id="1"/>' +
    p('between') +
    `<w:moveToRangeStart w:id="4" w:name="m"${WHO('Ann')}/>` +
    p(`<w:moveTo w:id="5"${WHO('Ann')}>${t('moved')}</w:moveTo>`, `<w:rPr><w:moveTo w:id="6"${WHO('Ann')}/></w:rPr>`) +
    '<w:moveToRangeEnd w:id="4"/>' +
    p('end');
  let o = await open(moved);
  assert.deepEqual(
    revisionsOf(o.doc).map((revision) => [revision.kind, revision.text, revision.move]),
    [
      ['moveFrom', 'moved¶', 'm'],
      ['moveTo', 'moved¶', 'm'],
    ],
  );
  // The paragraph that was moved away is empty and its mark goes: the next paragraph takes its place.
  assert.equal(await applied(o, acceptRevision(o.doc, first(o.doc))), `<w:p>${t('between')}</w:p><w:p>${t('moved')}</w:p>` + p('end'));
  o = await open(moved);
  assert.equal(await applied(o, rejectRevision(o.doc, first(o.doc, 'moveTo'))), `<w:p>${t('moved')}</w:p>` + p('between') + `<w:p>${t('end')}</w:p>`);

  const row = `<w:tr><w:trPr><w:ins w:id="1"${WHO('Ann')}/></w:trPr><w:tc>${p(t('new') + REF(2))}</w:tc></w:tr>`;
  o = await open(tbl(rowOf('old') + row, [2000]) + p('end'), { footnotes: NOTES(NOTE(2, 'two')) });
  run(o.history, rejectAll(o.doc));
  assert.deepEqual(texts(o.doc), ['old', 'end']);
  assert.deepEqual(Array.from(o.doc.footnotes!.items.keys()), [-1, 0]);
  o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);
});
