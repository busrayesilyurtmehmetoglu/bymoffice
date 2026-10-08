// What the independent review of accepting and rejecting found: every case
// here failed before its fix.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import type { DocModel, Paragraph, Table } from '../src/model/doc/types';
import {
  DocHistory,
  REVISION_GONE,
  REVISION_UNSUPPORTED,
  acceptAll,
  acceptRevision,
  deleteRange,
  insertText,
  listRevisions,
  rejectAll,
  rejectRevision,
  revisionCounters,
  revisionStep,
  revisionsOf,
  revisionsTouching,
  setTrackRevisions,
  setTracking,
  splitParagraph,
  type DocRange,
  type DocRevision,
} from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { parseXml } from '../src/xml';
import { STYLES, checkInvariants, p, pos, range, run, tbl, tc, texts } from './doc-ops-fixtures';
import { buildDocx as buildSpec, type DocxSpec } from './docx-fixtures';
import { schemaOrderViolations } from './docx-write-helpers.test';

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

async function body(o: Opened): Promise<string> {
  const text = await written(o);
  return text.slice(text.indexOf('<w:body>') + 8, text.lastIndexOf('<w:sectPr>'));
}

async function undone(o: Opened): Promise<void> {
  while (o.history.canUndo) o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes, 'undo gives the bytes that were read');
}

const WHO = (author: string, date = '2026-10-01T09:00:00Z') => ` w:author="${author}" w:date="${date}"`;
const ins = (id: number, inner: string, author = 'Ann', date?: string) => `<w:ins w:id="${id}"${WHO(author, date)}>${inner}</w:ins>`;
const del = (id: number, inner: string, author = 'Ann', date?: string) => `<w:del w:id="${id}"${WHO(author, date)}>${inner}</w:del>`;
const space = (text: string) => (/^\s|\s$/.test(text) ? ' xml:space="preserve"' : '');
const t = (text: string) => `<w:r><w:t${space(text)}>${text}</w:t></w:r>`;
const dt = (text: string) => `<w:r><w:delText${space(text)}>${text}</w:delText></w:r>`;
const markIns = (id: number, author = 'Ann') => `<w:rPr><w:ins w:id="${id}"${WHO(author)}/></w:rPr>`;
const markDel = (id: number, author = 'Ann') => `<w:rPr><w:del w:id="${id}"${WHO(author)}/></w:rPr>`;
const kinds = (doc: DocModel): string[] => revisionsOf(doc).map((revision) => revision.kind);
const first = (doc: DocModel, kind?: string): DocRevision => {
  const found = revisionsOf(doc).find((revision) => kind === undefined || revision.kind === kind);
  assert.ok(found, `no revision${kind ? ' of kind ' + kind : ''}`);
  return found;
};
const refused = (make: () => unknown, message: string): void => {
  assert.throws(make, (error: unknown) => error instanceof RangeError && error.message === message);
};
/** Every w:id of a tracked change in the text, with the element it is on. */
const changeIds = (xml: string): string[] => Array.from(xml.matchAll(/<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|sectPrChange|tblPrChange|trPrChange|tcPrChange)\b[^>]*? w:id="(\d+)"/g), (match) => match[1]);
const unique = (xml: string): void => {
  const ids = changeIds(xml);
  assert.equal(new Set(ids).size, ids.length, `an id twice: ${ids.join(',')}\n${xml}`);
};

void [checkInvariants, pos, range, tbl, tc, texts, kinds, markIns, markDel, del, dt, refused, unique, undone, body, first, run, REVISION_UNSUPPORTED, acceptAll, acceptRevision, deleteRange, insertText, listRevisions, rejectAll, setTrackRevisions, setTracking, splitParagraph];
export type { Paragraph, Table };

// ------------------------------------------------------------------------- B5

test('review B5: a copy of a description names the change at the place it describes, or none', async () => {
  const o = await open(p(ins(1, t('same')) + t(' a ') + ins(2, t('same')) + t(' b ') + ins(3, t('same'))));
  const copy = JSON.parse(JSON.stringify(revisionsOf(o.doc))) as DocRevision[];
  // The first one goes: the list shifts, and the copy of the second (index 1) now sits where another change is listed.
  run(o.history, rejectRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.deepEqual(revisionsOf(o.doc).map((revision) => revision.start.offset), [3, 10]);
  // Index 1 of the old list was the change at offset 7; index 1 of the new list is the one at offset 10.
  refused(() => rejectRevision(o.doc, copy[1]), REVISION_GONE);
  assert.deepEqual(texts(o.doc), [' a same b same']);
  // A copy that still describes what is there resolves exactly that.
  const fresh = JSON.parse(JSON.stringify(revisionsOf(o.doc))) as DocRevision[];
  run(o.history, rejectRevision(o.doc, fresh[1]));
  assert.deepEqual(texts(o.doc), [' a same b ']);
});

// ------------------------------------------------------------------------- B1

test('review B1: a move whose other end is in a text box (kept XML) is not resolved at one end alone', async () => {
  const box =
    '<w:r><w:pict><v:rect style="width:100pt;height:40pt"><v:textbox><w:txbxContent>' +
    `<w:p><w:moveToRangeStart w:id="3" w:name="a"${WHO('Ann')}/><w:moveTo w:id="4"${WHO('Ann')}>${t('X')}</w:moveTo><w:moveToRangeEnd w:id="3"/></w:p>` +
    '</w:txbxContent></v:textbox></v:rect></w:pict></w:r>';
  const o = await open(
    p(t('a ') + `<w:moveFromRangeStart w:id="1" w:name="a"${WHO('Ann')}/><w:moveFrom w:id="2"${WHO('Ann')}>${t('X')}</w:moveFrom><w:moveFromRangeEnd w:id="1"/>` + t(' b')) + p(box) + p(ins(9, t('new'), 'Bob')),
  );
  const list = revisionsOf(o.doc);
  assert.deepEqual(
    list.map((revision) => [revision.kind, revision.move, revision.canAccept, revision.canReject]),
    [
      ['moveFrom', 'a', false, false],
      ['other', undefined, false, false],
      ['insert', undefined, true, true],
    ],
  );
  refused(() => acceptRevision(o.doc, list[0]), REVISION_UNSUPPORTED);
  refused(() => rejectRevision(o.doc, list[0]), REVISION_UNSUPPORTED);
  for (const all of [acceptAll(o.doc), rejectAll(o.doc)]) assert.deepEqual([all.resolved, all.skipped], [1, 2]);
  run(o.history, rejectAll(o.doc));
  // Both ends are as they were.
  const out = await body(o);
  assert.ok(out.indexOf('<w:moveFromRangeStart w:id="1" w:name="a"') > 0 && out.indexOf(`<w:moveFrom w:id="2"${WHO('Ann')}>`) > 0 && out.indexOf('<w:moveToRangeStart w:id="3" w:name="a"') > 0, out);
  assert.deepEqual(kinds(o.doc), ['moveFrom', 'other']);
  await undone(o);
});

test('review: kept XML is looked through once, however often the changes are listed', async () => {
  const pict = (n: number) => `<w:r><w:pict><v:rect style="width:10pt;height:10pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>box ${n}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></w:r>`;
  let content = '';
  for (let i = 0; i < 20; i++) content += p(t('text ') + pict(i));
  const o = await open(content + p('end'));
  const before = revisionCounters.scanned;
  assert.deepEqual(revisionsOf(o.doc), []);
  const first = revisionCounters.scanned - before;
  assert.ok(first >= 20 * 5, String(first));
  revisionsOf(o.doc);
  assert.equal(revisionCounters.scanned - before, first, 'the second listing looks at no kept node again');
  // An edit makes new paragraphs; the drawings in them are the objects they were.
  run(o.history, insertText(o.doc, pos(3, 0), 'typed'));
  revisionsOf(o.doc);
  assert.equal(revisionCounters.scanned - before, first);
});

// ------------------------------------------------------------------------- B3

test('review B3: a field that runs over two paragraphs goes whole or not at all', async () => {
  const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>';
  const rest = '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + t('res') + '<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const code = '<w:r><w:delInstrText xml:space="preserve"> QUOTE x </w:delInstrText></w:r>';
  // The beginning and the code were deleted; the separator, the result and the end were not.
  let o = await open(p(t('a') + del(1, begin + code)) + p(rest + t('b')) + p(ins(2, t('new'), 'Bob')));
  refused(() => acceptRevision(o.doc, first(o.doc, 'delete')), REVISION_UNSUPPORTED);
  const all = acceptAll(o.doc);
  assert.deepEqual([all.resolved, all.skipped], [1, 1]);
  run(o.history, all);
  const out = await body(o);
  assert.ok(out.indexOf(del(1, begin + code)) > 0, out);
  assert.equal(out.split('fldCharType="begin"').length, 2);
  // Rejecting it takes no mark of the field away: that works.
  run(o.history, rejectRevision(o.doc, first(o.doc, 'delete')));
  assert.ok((await body(o)).indexOf('<w:instrText xml:space="preserve"> QUOTE x </w:instrText>') > 0);
  await undone(o);

  // All of the field deleted, over both paragraphs and the mark between them: it goes.
  o = await open(p(t('a') + del(1, begin + code), markDel(2)) + p(del(3, rest.replace('<w:t>res</w:t>', '<w:delText>res</w:delText>')) + t('b')));
  run(o.history, acceptAll(o.doc));
  assert.equal(await body(o), `<w:p>${t('ab')}</w:p>`);

  // The end of a field in a row that goes, its beginning outside of it.
  const row = `<w:tr><w:trPr><w:del w:id="5"${WHO('Ann')}/></w:trPr><w:tc>${p(rest)}</w:tc></w:tr>`;
  o = await open(p(t('a') + begin + '<w:r><w:instrText> QUOTE x </w:instrText></w:r>') + tbl(`<w:tr>${tc('keep')}</w:tr>` + row, [2000]) + p('end'));
  refused(() => acceptRevision(o.doc, first(o.doc)), REVISION_UNSUPPORTED);
  assert.deepEqual([acceptAll(o.doc).resolved, acceptAll(o.doc).skipped], [0, 1]);
});

// ------------------------------------------------------------------------- B2

test('review B2: joined paragraphs are formatted like the second one (Word), in whatever order the changes are resolved', async () => {
  const xml = p(ins(1, t('typed'), 'Cid'), '<w:jc w:val="center"/>' + markDel(2, 'Bob')) + p(t('second'), '<w:jc w:val="right"/>') + p('end');
  const outcome = async (order: string[]): Promise<string> => {
    const o = await open(xml);
    for (const author of order) {
      const revision = revisionsOf(o.doc).find((one) => one.author === author) as DocRevision;
      run(o.history, author === 'Bob' ? acceptRevision(o.doc, revision) : rejectRevision(o.doc, revision));
    }
    const out = await body(o);
    await undone(o);
    return out;
  };
  const expected = `<w:p><w:pPr><w:jc w:val="right"/></w:pPr>${t('second')}</w:p>` + p('end');
  assert.equal(await outcome(['Bob', 'Cid']), expected);
  assert.equal(await outcome(['Cid', 'Bob']), expected);

  // The paragraph that stays keeps what its own properties record: the change of its formatting is still Ann's to decide.
  const change = `<w:jc w:val="center"/><w:pPrChange w:id="5"${WHO('Ann')}><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange>`;
  const o = await open(p(t('one'), '<w:pStyle w:val="Heading1"/>' + markDel(1, 'Bob')) + p(t('two'), change) + p('end'));
  run(o.history, acceptRevision(o.doc, first(o.doc, 'delete')));
  assert.equal(await body(o), `<w:p><w:pPr>${change}</w:pPr>${t('onetwo')}</w:p>` + p('end'));
  assert.deepEqual(kinds(o.doc), ['paragraphFormat']);
});

test('review B2: a section break that goes with its paragraph mark is counted', async () => {
  const section = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';
  const o = await open(p(t('one'), section + markDel(1, 'Bob')) + p(t('two'), markIns(2, 'Ann')) + p('three') + p('end'));
  const one = acceptRevision(o.doc, first(o.doc, 'delete'));
  assert.equal(one.sectionsRemoved, 1);
  // The other mark ends no section.
  assert.equal(rejectRevision(o.doc, first(o.doc, 'insert')).sectionsRemoved, 0);
  const all = acceptAll(o.doc);
  assert.equal(all.sectionsRemoved, 1);
  run(o.history, all);
  const out = await written(o);
  assert.equal(out.split('<w:sectPr>').length, 2, out);
  assert.deepEqual(texts(o.doc), ['onetwo', 'three', 'end']);
  await undone(o);
});

// ------------------------------------------------------------------------- B4

const CHANGED = (id: number, text: string) => `<w:r><w:rPr><w:b/><w:rPrChange w:id="${id}"${WHO('Ann')}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>${text}</w:t></w:r>`;
const ME = { author: 'Me', date: '2026-10-08T10:00:00Z' };

test('review B4: text typed into a run whose formatting was changed does not take the record of that change along', async () => {
  for (const tracked of [false, true]) {
    const o = await open(p(CHANGED(5, 'abcd')), tracked ? { settings: '<w:trackRevisions/>' } : {});
    run(o.history, insertText(o.doc, pos(0, 2), 'X', undefined, ME));
    const out = await body(o);
    // The typed text is bold like its neighbours; the record of who changed their formatting stays theirs.
    assert.ok(/<w:r><w:rPr><w:b\/><\/w:rPr><w:t>X<\/w:t><\/w:r>/.test(out), out);
    assert.equal(out.split('<w:rPrChange').length, 3, out);
    unique(out);
    assert.deepEqual(
      revisionsOf(o.doc).map((revision) => [revision.kind, revision.text]),
      tracked
        ? [
            ['format', 'ab'],
            ['insert', 'X'],
            ['format', 'cd'],
          ]
        : [
            ['format', 'ab'],
            ['format', 'cd'],
          ],
    );
    await undone(o);
  }
});

test('review B4: what an edit splits in two has two ids, with and without tracking', async () => {
  const cases: [name: string, xml: string][] = [
    ['run with a formatting change', p(CHANGED(5, 'abcdef'))],
    ['insertion', p(ins(5, t('abcdef'), 'Ann'))],
    ['move', p(`<w:moveTo w:id="5"${WHO('Ann')}>${t('abcdef')}</w:moveTo>`)],
    ['paragraph with a formatting change', p(t('abcdef'), `<w:jc w:val="center"/><w:rPr><w:b/><w:rPrChange w:id="6"${WHO('Ann')}><w:rPr/></w:rPrChange></w:rPr><w:pPrChange w:id="5"${WHO('Ann')}><w:pPr/></w:pPrChange>`)],
  ];
  for (const tracked of [false, true]) {
    for (const [name, xml] of cases) {
      const edits: [string, (o: Opened) => void][] = [
        ['typing in the middle', (o) => run(o.history, insertText(o.doc, pos(0, 3), 'X', undefined, ME))],
        ['deleting from the middle', (o) => run(o.history, deleteRange(o.doc, range(pos(0, 2), pos(0, 4)), ME))],
        ['splitting the paragraph', (o) => run(o.history, splitParagraph(o.doc, pos(0, 3), ME))],
        ['typing a line end', (o) => run(o.history, insertText(o.doc, pos(0, 3), 'X\nY', undefined, ME))],
      ];
      for (const [what, edit] of edits) {
        const o = await open(xml + p('end'), tracked ? { settings: '<w:trackRevisions/>' } : {});
        edit(o);
        checkInvariants(o.doc);
        unique(await body(o));
        // What is in memory is what the file says: listed again from the file, the changes are the same.
        const again = await readDocx(await OpcPackage.open(await o.pkg.save()));
        assert.deepEqual(
          revisionsOf(again).map((revision) => [revision.kind, revision.author, revision.text]),
          revisionsOf(o.doc).map((revision) => [revision.kind, revision.author, revision.text]),
          `${tracked ? 'tracked' : 'plain'}: ${name}: ${what}`,
        );
        await undone(o);
      }
    }
  }
});

test('review B4: one wrapper object around two stretches with plain text between them is listed as two changes', async () => {
  const o = await open(p(ins(1, t('one')) + t(' plain ') + ins(2, t('two'))));
  // A model as an edit that does not renumber would leave it: both stretches in the same wrapper object.
  const paragraph = o.doc.body.blocks[0] as Paragraph;
  const [a, plain, b] = paragraph.children;
  o.doc.body.blocks[0] = { ...paragraph, children: [a, plain, { ...b, wrap: a.wrap } as never] };
  assert.deepEqual(
    revisionsOf(o.doc).map((revision) => [revision.kind, revision.text, revision.start.offset]),
    [
      ['insert', 'one', 0],
      ['insert', 'two', 10],
    ],
  );
});

// ------------------------------------------------- branches no test had held

test('review: "all" with a filter that names one end of a move resolves the other end with it', async () => {
  const move =
    p(t('a ') + `<w:moveFromRangeStart w:id="1" w:name="m"${WHO('Ann')}/><w:moveFrom w:id="2"${WHO('Ann')}>${t('X')}</w:moveFrom><w:moveFromRangeEnd w:id="1"/>` + t(' b')) +
    p(t('c ') + `<w:moveToRangeStart w:id="3" w:name="m"${WHO('Ann')}/><w:moveTo w:id="4"${WHO('Ann')}>${t('X')}</w:moveTo><w:moveToRangeEnd w:id="3"/>` + t(' d'));
  const o = await open(move + p(ins(9, t('other'), 'Bob')));
  const all = rejectAll(o.doc, (revision) => revision.kind === 'moveFrom');
  assert.deepEqual([all.resolved, all.skipped], [2, 0]);
  run(o.history, all);
  assert.deepEqual(texts(o.doc), ['a X b', 'c  d', 'other']);
  assert.deepEqual(kinds(o.doc), ['insert']);
});

test('review: what a rejected formatting change does not record stays as it is', async () => {
  // The section a paragraph ends is not in <w:pPrChange>.
  const section = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>';
  let o = await open(p(t('one'), `<w:jc w:val="center"/>${section}<w:pPrChange w:id="1"${WHO('Ann')}><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange>`) + p('end'));
  run(o.history, rejectRevision(o.doc, first(o.doc)));
  assert.equal(await body(o), `<w:p><w:pPr><w:jc w:val="left"/>${section}</w:pPr>${t('one')}</w:p>` + p('end'));

  // The headers and footers of a section are not in <w:sectPrChange>.
  const last = `<w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:pgSz w:w="8000" w:h="8000"/><w:sectPrChange w:id="2"${WHO('Ann')}><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:sectPrChange></w:sectPr>`;
  const bytes = await buildSpec({ body: p('one') + last, styles: STYLES, headers: { rId7: p('head') } });
  const pkg = await OpcPackage.open(bytes);
  o = { pkg, doc: await readDocx(pkg), history: new DocHistory<DocRange | undefined>(await readDocx(pkg)), bytes };
  o.history = new DocHistory<DocRange | undefined>(o.doc);
  run(o.history, rejectRevision(o.doc, first(o.doc)));
  const out = await written(o);
  assert.ok(out.indexOf('<w:sectPr><w:headerReference w:type="default" r:id="rId7"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>') > 0, out);

  // That a row was inserted is not in <w:trPrChange>.
  const row = `<w:tr><w:trPr><w:cantSplit/><w:ins w:id="3"${WHO('Bob')}/><w:trPrChange w:id="4"${WHO('Ann')}><w:trPr/></w:trPrChange></w:trPr>${tc('cell')}</w:tr>`;
  o = await open(tbl(row, [2000]) + p('end'));
  run(o.history, rejectRevision(o.doc, first(o.doc, 'tableFormat')));
  let text = await body(o);
  assert.ok(text.indexOf(`<w:trPr><w:ins w:id="3"${WHO('Bob')}/></w:trPr>`) > 0, text);
  assert.deepEqual(kinds(o.doc), ['insert']);

  // That a paragraph mark was inserted is not in the <w:rPrChange> of the mark.
  o = await open(p(t('one'), `<w:rPr><w:ins w:id="5"${WHO('Bob')}/><w:b/><w:rPrChange w:id="6"${WHO('Ann')}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr>`) + p('end'));
  run(o.history, rejectRevision(o.doc, first(o.doc, 'format')));
  text = await body(o);
  assert.ok(text.indexOf(`<w:pPr><w:rPr><w:ins w:id="5"${WHO('Bob')}/><w:i/></w:rPr></w:pPr>`) > 0, text);
});

test('review: the marks of bookmarks and comments of a row that goes stay, in the row beside it', async () => {
  const marks = '<w:bookmarkStart w:id="20" w:name="kept"/>' + t('gone') + '<w:bookmarkEnd w:id="20"/><w:commentRangeStart w:id="4"/><w:commentRangeEnd w:id="4"/><w:r><w:commentReference w:id="4"/></w:r>';
  const row = `<w:tr><w:trPr><w:ins w:id="1"${WHO('Ann')}/></w:trPr><w:tc>${p(marks)}</w:tc></w:tr>`;
  const o = await open(tbl(`<w:tr>${tc('old')}</w:tr>` + row, [2000]) + p('end'), { comments: `<w:comment w:id="4" w:author="Ann"><w:p>${t('note')}</w:p></w:comment>` });
  run(o.history, rejectAll(o.doc));
  const out = await body(o);
  assert.deepEqual(texts(o.doc), ['old', 'end']);
  for (const kept of ['<w:bookmarkStart w:id="20" w:name="kept"/>', '<w:bookmarkEnd w:id="20"/>', '<w:commentRangeStart w:id="4"/>', '<w:commentRangeEnd w:id="4"/>', '<w:commentReference w:id="4"/>']) assert.ok(out.indexOf(kept) > 0, `${kept}\n${out}`);
  assert.ok(out.indexOf('gone') < 0, out);
  await undone(o);
});

test('review: a copy of a description that says another text names no change', async () => {
  const o = await open(p(ins(1, t('same')) + t(' a ') + ins(2, t('other'))));
  const [copy] = JSON.parse(JSON.stringify(revisionsOf(o.doc))) as DocRevision[];
  refused(() => acceptRevision(o.doc, { ...copy, text: 'else' }), REVISION_GONE);
  refused(() => acceptRevision(o.doc, { ...copy, end: { ...copy.end, offset: 3 } }), REVISION_GONE);
  run(o.history, acceptRevision(o.doc, copy));
  assert.deepEqual(kinds(o.doc), ['insert']);
});

// ------------------------------------------------- B6 and the caret on a border

test('review B6: the changes beyond the limit of the list are found at the caret and stepped to', async () => {
  let content = '';
  for (let k = 0; k < 41; k++) {
    let inner = '';
    for (let i = 0; i < 50; i++) inner += ins(k * 50 + i + 1, t('x'), 'A' + ((k * 50 + i) % 7)) + t('.');
    content += p(inner);
  }
  const o = await open(content + p('end'));
  const listed = listRevisions(o.doc);
  assert.deepEqual([listed.list.length, listed.more], [2000, true]);
  const all = listRevisions(o.doc, { limit: Infinity });
  assert.deepEqual([all.list.length, all.more], [2050, false]);
  // The caret in change 2031 (paragraph 40, the 31st insertion there, at offsets 60..61).
  const caret = { anchor: pos(40, 61), focus: pos(40, 61) };
  assert.deepEqual(revisionsTouching(all.list, caret), [2030]);
  assert.deepEqual(revisionsTouching(listed.list, caret), []);
  // From the last one of the bounded list the step goes on, not round.
  assert.deepEqual(revisionStep(all.list, undefined, 'next', 1999), { index: 2000 });
  assert.deepEqual(revisionStep(all.list, undefined, 'next', 2049), { index: 0, wrapped: 'first' });
  assert.deepEqual(revisionStep(all.list, undefined, 'previous', 0), { index: 2049, wrapped: 'last' });
  // Without a change stepped to before, from where the selection is.
  assert.deepEqual(revisionStep(all.list, caret, 'next', -1), { index: 2031 });
  assert.deepEqual(revisionStep(all.list, { anchor: pos(40, 60), focus: pos(40, 60) }, 'next', -1), { index: 2030 });
  assert.deepEqual(revisionStep(all.list, caret, 'previous', -1), { index: 2030 });
  assert.deepEqual(revisionStep(all.list, { anchor: pos(41, 1), focus: pos(41, 1) }, 'next', -1), { index: 0, wrapped: 'first' });
  assert.equal(revisionStep([], caret, 'next', -1), undefined);
  // A change found there is resolved like any other.
  run(o.history, rejectRevision(o.doc, all.list[2030]));
  assert.equal(listRevisions(o.doc, { limit: Infinity }).list.length, 2049);
});

test('review: a caret on the border of two changes is in one of them: the one it is inside, else the first', async () => {
  const o = await open(p(t('One ') + ins(1, t('two ')) + del(2, dt('three '), 'Bob') + ins(3, t('four'), 'Cid') + t('.')));
  const list = revisionsOf(o.doc);
  const at = (offset: number, to = offset) => revisionsTouching(list, { anchor: pos(0, offset), focus: pos(0, to) });
  assert.deepEqual(at(6), [0]);
  // At offset 8 the first insertion ends, the deletion stands and the second insertion begins.
  assert.deepEqual(at(8), [0]);
  assert.deepEqual(at(4), [0]);
  assert.deepEqual(at(10), [2]);
  assert.deepEqual(at(12), [2]);
  assert.deepEqual(at(2), []);
  // A selection holds every change it runs over, and a deletion that stands inside it.
  assert.deepEqual(at(5, 10), [0, 1, 2]);
  assert.deepEqual(at(0, 8), [0]);
  assert.deepEqual(at(8, 13), [2]);
  assert.deepEqual(at(0, 4), []);
});
