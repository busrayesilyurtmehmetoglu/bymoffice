// Formatting in a document that tracks its changes: character formatting is
// recorded on the runs it changes (<w:rPrChange>), paragraph formatting and
// lists on the paragraphs (<w:pPrChange>), each with the formatting as it was;
// rejecting brings that back, accepting drops the record.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { buildDocx, writeDocx } from '../src/formats/docx/write';
import { readRtf } from '../src/formats/rtf/document';
import type { DocModel, Paragraph, Run } from '../src/model/doc/types';
import {
  DocHistory,
  TRACK_UNSUPPORTED,
  acceptAll,
  acceptRevision,
  deleteBackward,
  deleteColumn,
  deleteForward,
  deleteRange,
  indentList,
  insertBlock,
  insertColumn,
  insertText,
  listRevisions,
  rejectAll,
  rejectRevision,
  revisionsOf,
  setParaFormat,
  setRunFormat,
  setTracking,
  splitParagraph,
  storyIds,
  toggleList,
  toggleRunFormat,
  type DocEdit,
  type DocEditOptions,
  type DocPosition,
  type DocRange,
} from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { parseXml, type XmlElement } from '../src/xml';
import { NUMBERING, STYLES, checkInvariants, dice, docPicture, p, paragraphAt, paragraphPaths, picture, pos, r, randomPosition, range, run, tbl, tc, texts, tr } from './doc-ops-fixtures';
import { buildDocx as buildSpec, type DocxSpec } from './docx-fixtures';
import { schemaOrderViolations } from './docx-write-helpers.test';

const decoder = new TextDecoder();
const DATE = '2026-10-05T10:00:00Z';
const ME: DocEditOptions = { author: 'Me', date: DATE };
const YOU: DocEditOptions = { author: 'You', date: '2026-10-06T10:00:00Z' };
const TRACK = '<w:trackRevisions/>';
const SECT = '<w:sectPr><w:pgSz w:w="9000" w:h="9000"/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr>';

interface Opened {
  pkg: OpcPackage;
  doc: DocModel;
  history: DocHistory<DocRange | undefined>;
  bytes: Uint8Array;
}

async function open(body: string, extra: Partial<DocxSpec> = {}): Promise<Opened> {
  const bytes = await buildSpec({ body: body + SECT, styles: STYLES, numbering: NUMBERING, settings: TRACK, ...extra });
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  return { pkg, doc, history: new DocHistory<DocRange | undefined>(doc), bytes };
}

/** The document part as the writer writes it now. */
async function written(o: Opened, part = '/word/document.xml'): Promise<string> {
  await writeDocx(o.pkg, o.doc);
  return decoder.decode(await o.pkg.readPart(part));
}

/** The body of the written document part, without the section. */
async function body(o: Opened): Promise<string> {
  const text = await written(o);
  return text.slice(text.indexOf('<w:body>') + 8, text.lastIndexOf('<w:sectPr>'));
}

async function reopened(o: Opened): Promise<DocModel> {
  await writeDocx(o.pkg, o.doc);
  return readDocx(await OpcPackage.open(await o.pkg.save()));
}

const WHO = ` w:author="Me" w:date="${DATE}"`;
const t = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t>${text}</w:t></w:r>`;
/** A record of changed character formatting, with what the formatting was. */
const was = (id: number, old = '', who = WHO) => `<w:rPrChange w:id="${id}"${who}>${old ? `<w:rPr>${old}</w:rPr>` : '<w:rPr/>'}</w:rPrChange>`;
/** A record of changed paragraph formatting. */
const pWas = (id: number, old = '', who = WHO) => `<w:pPrChange w:id="${id}"${who}>${old ? `<w:pPr>${old}</w:pPr>` : '<w:pPr/>'}</w:pPrChange>`;

const refused = (make: () => DocEdit): void => {
  assert.throws(make, (error: unknown) => error instanceof RangeError && error.message === TRACK_UNSUPPORTED);
};

/** The tracked changes as [kind, author, text]. */
const listed = (doc: DocModel): [string, string | undefined, string][] => revisionsOf(doc).map((revision) => [revision.kind, revision.author, revision.text]);

/** Every w:id of a tracked change in the XML. */
function revisionIds(xml: string): string[] {
  const out: string[] = [];
  for (const match of xml.matchAll(/<w:(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|sectPrChange|tblPrChange|trPrChange|tcPrChange)\b[^>]*? w:id="(\d+)"/g)) out.push(match[1]);
  return out;
}

const local = (name: string): string => name.slice(name.indexOf(':') + 1);
const elements = (el: XmlElement): XmlElement[] => el.children.filter((node): node is XmlElement => node.type === 'element');

/**
 * Where the schema wants the records of changed formatting: <w:rPrChange> the last child of an <w:rPr>, holding one
 * <w:rPr> and nothing else, which holds no record itself; <w:pPrChange> the last child of a <w:pPr>, holding one
 * <w:pPr> without <w:rPr>, <w:sectPr> or <w:pPrChange>. Every record has an id, an author and a date.
 */
function misplacedRecords(xml: string): string[] {
  const found: string[] = [];
  const walk = (el: XmlElement, parent: XmlElement | undefined, depth: number): void => {
    if (depth > 400) return;
    const name = local(el.name);
    if (name === 'rPrChange' || name === 'pPrChange') {
      const outer = name === 'rPrChange' ? 'rPr' : 'pPr';
      const siblings = parent ? elements(parent) : [];
      if (!parent || local(parent.name) !== outer) found.push(`${name} in ${parent ? parent.name : 'nothing'}`);
      else if (siblings[siblings.length - 1] !== el) found.push(`${name} is not the last child`);
      const inner = elements(el);
      if (inner.length !== 1 || local(inner[0].name) !== outer) found.push(`${name} does not hold one ${outer}`);
      else {
        const banned = name === 'rPrChange' ? ['rPrChange', 'ins', 'del', 'moveFrom', 'moveTo'] : ['rPr', 'sectPr', 'pPrChange'];
        for (const child of elements(inner[0])) if (banned.indexOf(local(child.name)) >= 0) found.push(`${name} holds ${child.name}`);
      }
      for (const attr of ['w:id', 'w:author', 'w:date']) if (!el.attrs.has(attr)) found.push(`${name} without ${attr}`);
    }
    for (const child of elements(el)) walk(child, el, depth + 1);
  };
  walk(parseXml(xml).root, undefined, 0);
  return found;
}

function checkWritten(xml: string, label = ''): void {
  parseXml(xml);
  assert.deepEqual(schemaOrderViolations(xml), [], label);
  assert.deepEqual(misplacedRecords(xml), [], label);
  const ids = revisionIds(xml);
  assert.equal(new Set(ids).size, ids.length, `${label} ids: ${ids.join(' ')}`);
}

// ------------------------------------------------------------- character formatting

test('bold over part of a run is recorded on that part, with the formatting it had', async () => {
  const o = await open(p('abcd') + p(r('efgh', '<w:i/>')));
  run(o.history, setRunFormat(o.doc, range(pos(0, 1), pos(0, 3)), { bold: true }, ME));
  assert.equal(await body(o), `<w:p>${t('a')}${t('bc', '<w:b/><w:bCs/>' + was(1))}${t('d')}</w:p><w:p>${r('efgh', '<w:i/>')}</w:p>`);
  run(o.history, setRunFormat(o.doc, range(pos(1, 0), pos(1, 2)), { size: 40, color: 'FF0000' }, ME));
  assert.equal(
    await body(o),
    `<w:p>${t('a')}${t('bc', '<w:b/><w:bCs/>' + was(1))}${t('d')}</w:p><w:p>${t('ef', '<w:i/><w:color w:val="FF0000"/><w:sz w:val="40"/><w:szCs w:val="40"/>' + was(2, '<w:i/>'))}${t('gh', '<w:i/>')}</w:p>`,
  );
  assert.deepEqual(listed(o.doc), [
    ['format', 'Me', 'bc'],
    ['format', 'Me', 'ef'],
  ]);
  checkWritten(await written(o));
  const again = await reopened(o);
  assert.deepEqual(listed(again), listed(o.doc));
  // The model that is read is the model that was written (where in the source a paragraph stood aside).
  for (const index of [0, 1]) {
    assert.equal(picture(paragraphAt(again, index).children), picture(paragraphAt(o.doc, index).children), `paragraph ${index}`);
    assert.equal(picture(paragraphAt(again, index).props), picture(paragraphAt(o.doc, index).props), `paragraph ${index}`);
  }
  checkInvariants(o.doc, false);
});

test('every kind of character formatting the editor applies is tracked', async () => {
  const patches = [{ italic: true }, { underline: true }, { strike: true }, { font: 'Other' }, { size: 30 }, { color: '00FF00' }, { highlight: 'yellow' }, { vertAlign: 'superscript' as const }, { style: 'Strong' }];
  for (const patch of patches) {
    const o = await open(p('abcd'));
    const before = formatted(o.doc);
    run(o.history, setRunFormat(o.doc, range(pos(0, 0), pos(0, 4)), patch, ME));
    const label = JSON.stringify(patch);
    assert.deepEqual(listed(o.doc), [['format', 'Me', 'abcd']], label);
    const xml = await written(o);
    checkWritten(xml, label);
    assert.ok(xml.indexOf(was(1)) > 0, `${label}: ${xml}`);
    run(o.history, rejectAll(o.doc));
    assert.equal(formatted(o.doc), before, label);
  }
  // The toggle of a property a style turns on: switched off on the text, and that is a change like any other.
  const o = await open(p('abcd', '<w:pStyle w:val="Heading1"/>'));
  run(o.history, toggleRunFormat(o.doc, range(pos(0, 0), pos(0, 2)), 'bold', ME));
  assert.equal(await body(o), `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${t('ab', '<w:b w:val="0"/><w:bCs w:val="0"/>' + was(1))}${t('cd')}</w:p>`);
});

test('a run whose formatting somebody changed already keeps that record and gets no second one', async () => {
  const theirs = '<w:rPrChange w:id="7" w:author="You" w:date="2020-01-01T00:00:00Z"><w:rPr><w:i/></w:rPr></w:rPrChange>';
  const o = await open(p(r('abcd', '<w:b/>' + theirs)));
  run(o.history, setRunFormat(o.doc, range(pos(0, 0), pos(0, 4)), { underline: true }, ME));
  // The first "as it was" is the one that counts.
  assert.equal(await body(o), `<w:p>${t('abcd', '<w:b/><w:u w:val="single"/>' + theirs)}</w:p>`);
  assert.deepEqual(listed(o.doc), [['format', 'You', 'abcd']]);
  // Rejecting it goes back to what the record says, not to what was there before the second change.
  run(o.history, rejectRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.equal(await body(o), `<w:p>${t('abcd', '<w:i/>')}</w:p>`);
});

test('a selection that cuts a run with a record leaves the record on both parts, each with an id of its own', async () => {
  const theirs = (id: number) => `<w:rPrChange w:id="${id}" w:author="You" w:date="2020-01-01T00:00:00Z"><w:rPr><w:i/></w:rPr></w:rPrChange>`;
  const o = await open(p(r('abcd', '<w:b/>' + theirs(7))) + p(`<w:ins w:id="9" w:author="You" w:date="2020-01-01T00:00:00Z">${r('wxyz')}</w:ins>`));
  run(o.history, setRunFormat(o.doc, range(pos(0, 1), pos(0, 3)), { underline: true }, ME));
  run(o.history, setRunFormat(o.doc, range(pos(1, 1), pos(1, 3)), { bold: true }, ME));
  const xml = await body(o);
  assert.equal(
    xml,
    `<w:p>${t('a', '<w:b/>' + theirs(7))}${t('bc', '<w:b/><w:u w:val="single"/>' + theirs(10))}${t('d', '<w:b/>' + theirs(11))}</w:p>` +
      `<w:p><w:ins w:id="9" w:author="You" w:date="2020-01-01T00:00:00Z">${t('w')}${t('xy', '<w:b/><w:bCs/>' + was(12))}${t('z')}</w:ins></w:p>`,
  );
  checkWritten(await written(o));
});

test('the author who brings the formatting back to what it was takes the record away; nobody else does', async () => {
  const o = await open(p('abcd'));
  const span = range(pos(0, 1), pos(0, 3));
  const original = await body(o);
  run(o.history, toggleRunFormat(o.doc, span, 'bold', ME));
  assert.equal(revisionsOf(o.doc).length, 1);
  run(o.history, toggleRunFormat(o.doc, span, 'bold', ME));
  assert.deepEqual(listed(o.doc), []);
  assert.equal(await body(o), `<w:p>${t('abcd')}</w:p>`, original);

  // Another author changing it back leaves the record of the first: it is not theirs.
  run(o.history, toggleRunFormat(o.doc, span, 'bold', ME));
  run(o.history, toggleRunFormat(o.doc, span, 'bold', YOU));
  assert.deepEqual(listed(o.doc), [['format', 'Me', 'bc']]);
  assert.equal(await body(o), `<w:p>${t('a')}${t('bc', was(2))}${t('d')}</w:p>`);

  // A result that is not what the record holds keeps the record.
  const other = await open(p('abcd'));
  run(other.history, setRunFormat(other.doc, span, { bold: true }, ME));
  run(other.history, setRunFormat(other.doc, span, { bold: false }, ME));
  assert.equal(await body(other), `<w:p>${t('a')}${t('bc', '<w:b w:val="0"/><w:bCs w:val="0"/>' + was(1))}${t('d')}</w:p>`);
  run(other.history, setRunFormat(other.doc, span, { bold: null }, ME));
  assert.deepEqual(listed(other.doc), []);

  // An editor without a name owns nothing, and a record that came with the file is nobody's of this session.
  const anonymous = await open(p('abcd'));
  setTracking(anonymous.doc, { now: () => new Date(0) });
  run(anonymous.history, toggleRunFormat(anonymous.doc, span, 'bold'));
  run(anonymous.history, toggleRunFormat(anonymous.doc, span, 'bold'));
  assert.deepEqual(listed(anonymous.doc), [['format', 'Author', 'bc']]);
  const fromFile = await open(p(r('abcd', `<w:b/><w:rPrChange w:id="3"${WHO}><w:rPr/></w:rPrChange>`)));
  run(fromFile.history, setRunFormat(fromFile.doc, range(pos(0, 0), pos(0, 4)), { bold: null }, ME));
  assert.deepEqual(listed(fromFile.doc), [['format', 'Me', 'abcd']]);
});

test('text the author inserted as a tracked change is formatted without a record; text of somebody else\'s insertion gets one', async () => {
  const o = await open(p('abcd'));
  run(o.history, insertText(o.doc, pos(0, 2), 'XYZ', undefined, ME));
  run(o.history, setRunFormat(o.doc, range(pos(0, 1), pos(0, 6)), { bold: true }, ME));
  const b = '<w:b/><w:bCs/>';
  assert.equal(await body(o), `<w:p>${t('a')}${t('b', b + was(2))}<w:ins w:id="1"${WHO}>${t('XYZ', b)}</w:ins>${t('c', b + was(3))}${t('d')}</w:p>`);
  // Somebody else formatting the same text: it is not their insertion.
  run(o.history, setRunFormat(o.doc, range(pos(0, 2), pos(0, 5)), { italic: true }, YOU));
  const xml = await body(o);
  assert.ok(xml.indexOf(`<w:ins w:id="1"${WHO}>${t('XYZ', b + `<w:i/><w:iCs/><w:rPrChange w:id="4" w:author="You" w:date="2026-10-06T10:00:00Z"><w:rPr>${b}</w:rPr></w:rPrChange>`)}</w:ins>`) > 0, xml);
  checkWritten(await written(o));
});

test('deleted text is passed over: its formatting stays what it was', async () => {
  const o = await open(p(r('ab') + `<w:del w:id="4" w:author="You" w:date="2020-01-01T00:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del>` + r('cd')));
  run(o.history, deleteRange(o.doc, range(pos(0, 3), pos(0, 4)), ME));
  run(o.history, setRunFormat(o.doc, range(pos(0, 0), pos(0, 3)), { bold: true }, ME));
  const b = '<w:b/><w:bCs/>';
  assert.equal(
    await body(o),
    `<w:p>${t('ab', b + was(6))}<w:del w:id="4" w:author="You" w:date="2020-01-01T00:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del>${t('c', b + was(7))}<w:del w:id="5"${WHO}><w:r><w:delText>d</w:delText></w:r></w:del></w:p>`,
  );
  checkWritten(await written(o));
});

test('a paragraph mark in the selection is formatted with a record of its own, unless the paragraph is all the author\'s', async () => {
  const o = await open(p('ab') + p('cd'));
  run(o.history, setRunFormat(o.doc, range(pos(0, 1), pos(1, 1)), { bold: true }, ME));
  const b = '<w:b/><w:bCs/>';
  assert.equal(await body(o), `<w:p><w:pPr><w:rPr>${b}${was(2)}</w:rPr></w:pPr>${t('a')}${t('b', b + was(1))}</w:p><w:p>${t('c', b + was(3))}${t('d')}</w:p>`);
  assert.deepEqual(
    revisionsOf(o.doc).map((revision) => revision.kind),
    // The end of the first paragraph and its mark are one change; the second paragraph begins another.
    ['format', 'format'],
  );
  checkWritten(await written(o));
  run(o.history, rejectAll(o.doc));
  assert.equal(await body(o), `<w:p>${t('ab')}</w:p><w:p>${t('cd')}</w:p>`);

  // A mark the author made, of a paragraph that is all theirs: its formatting is new with it.
  const own = await open(p('abcd'));
  run(own.history, insertText(own.doc, pos(0, 4), '\nxy\n', undefined, ME));
  run(own.history, setRunFormat(own.doc, range(pos(1, 2), pos(2, 0)), { bold: true }, ME));
  assert.equal((await body(own)).indexOf('rPrChange'), -1, await body(own));
  assert.ok((await body(own)).indexOf(`${b}</w:rPr></w:pPr><w:ins `) > 0, await body(own));

  // At a caret in an empty paragraph the formatting goes onto the mark, as a tracked change.
  const empty = await open(p('') + p('x'));
  run(empty.history, setRunFormat(empty.doc, range(pos(0, 0), pos(0, 0)), { italic: true }, ME));
  assert.equal(await body(empty), `<w:p><w:pPr><w:rPr><w:i/><w:iCs/>${was(1)}</w:rPr></w:pPr></w:p><w:p>${r('x')}</w:p>`);
});

test('runs that share their formatting get records with ids of their own, in one paragraph and across paragraphs', async () => {
  const o = await open(p(r('ab', '<w:i/>') + r('-') + r('cd', '<w:i/>')) + p(r('ef', '<w:i/>')) + p(r('gh', '<w:i/>')));
  run(o.history, setRunFormat(o.doc, range(pos(0, 0), pos(2, 2)), { bold: true }, ME));
  const xml = await written(o);
  checkWritten(xml);
  assert.equal(revisionIds(xml).length, 7, xml);
});

// ------------------------------------------------------------- paragraph formatting

test('paragraph formatting is recorded on the paragraph, with the properties it had and without its mark or its section', async () => {
  const section = '<w:sectPr><w:pgSz w:w="8000" w:h="8000"/></w:sectPr>';
  const o = await open(p('ab', `<w:spacing w:before="120"/><w:rPr><w:b/></w:rPr>${section}`) + p('cd') + p('ef'));
  run(o.history, setParaFormat(o.doc, range(pos(0, 1), pos(1, 1)), { align: 'center' }, ME));
  assert.equal(
    await body(o),
    `<w:p><w:pPr><w:spacing w:before="120"/><w:jc w:val="center"/><w:rPr><w:b/></w:rPr>${section}${pWas(1, '<w:spacing w:before="120"/>')}</w:pPr>${t('ab')}</w:p>` +
      `<w:p><w:pPr><w:jc w:val="center"/>${pWas(2)}</w:pPr>${t('cd')}</w:p><w:p>${r('ef')}</w:p>`,
  );
  assert.deepEqual(listed(o.doc), [
    ['paragraphFormat', 'Me', ''],
    ['paragraphFormat', 'Me', ''],
  ]);
  checkWritten(await written(o));
  const again = await reopened(o);
  assert.deepEqual(listed(again), listed(o.doc));
  for (const index of [0, 1, 2]) assert.equal(picture(paragraphAt(again, index).props), picture(paragraphAt(o.doc, index).props), `paragraph ${index}`);
});

test('indents, spacing, style, lists and levels are all changes of the paragraph', async () => {
  const cases: [string, (doc: DocModel, at: DocRange) => DocEdit, string][] = [
    ['indent', (doc, at) => setParaFormat(doc, at, { indent: { left: 720 } }, ME), '<w:ind w:left="720"/>'],
    ['spacing', (doc, at) => setParaFormat(doc, at, { spacing: { after: 200, line: 360, lineRule: 'auto' } }, ME), '<w:spacing w:after="200" w:line="360" w:lineRule="auto"/>'],
    ['style', (doc, at) => setParaFormat(doc, at, { style: 'Heading1' }, ME), '<w:pStyle w:val="Heading1"/>'],
    ['list', (doc, at) => toggleList(doc, at, 'bullet', ME), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="'],
    ['indent step', (doc, at) => indentList(doc, at, 1, ME), '<w:ind w:left="720"/>'],
  ];
  for (const [label, make, expected] of cases) {
    const o = await open(p('ab') + p('cd'));
    const before = formatted(o.doc);
    run(o.history, make(o.doc, range(pos(0, 0), pos(0, 1))));
    const xml = await body(o);
    assert.ok(xml.indexOf(`<w:p><w:pPr>${expected}`) === 0, `${label}: ${xml}`);
    assert.ok(xml.indexOf(`${pWas(1)}</w:pPr>${t('ab')}</w:p><w:p>${r('cd')}</w:p>`) > 0, `${label}: ${xml}`);
    assert.deepEqual(listed(o.doc), [['paragraphFormat', 'Me', '']], label);
    checkWritten(await written(o), label);
    run(o.history, rejectRevision(o.doc, revisionsOf(o.doc)[0]));
    assert.equal(formatted(o.doc), before, label);
  }
  // A list item taken out of its list, and one moved a level deeper: the record holds the list it was in.
  const numPr = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>';
  const out = await open(p('ab', numPr) + p('cd', numPr));
  run(out.history, toggleList(out.doc, range(pos(0, 0), pos(0, 1)), 'bullet', ME));
  run(out.history, indentList(out.doc, range(pos(1, 0), pos(1, 1)), 1, ME));
  assert.equal(await body(out), `<w:p><w:pPr>${pWas(1, numPr)}</w:pPr>${t('ab')}</w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>${pWas(2, numPr)}</w:pPr>${t('cd')}</w:p>`);
  checkWritten(await written(out));
});

test('a paragraph with a record keeps it; its author takes it away by going back; a paragraph that is all the author\'s needs none', async () => {
  const theirs = '<w:pPrChange w:id="7" w:author="You" w:date="2020-01-01T00:00:00Z"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange>';
  const o = await open(p('ab', '<w:jc w:val="center"/>' + theirs) + p('cd'));
  run(o.history, setParaFormat(o.doc, pos(0, 0), { align: 'both' }, ME));
  assert.equal(await body(o), `<w:p><w:pPr><w:jc w:val="both"/>${theirs}</w:pPr>${t('ab')}</w:p><w:p>${r('cd')}</w:p>`);
  // Going back to what their record says does not take their record away.
  run(o.history, setParaFormat(o.doc, pos(0, 0), { align: 'right' }, ME));
  assert.equal(await body(o), `<w:p><w:pPr><w:jc w:val="right"/>${theirs}</w:pPr>${t('ab')}</w:p><w:p>${r('cd')}</w:p>`);

  run(o.history, setParaFormat(o.doc, pos(1, 0), { align: 'center' }, ME));
  run(o.history, setParaFormat(o.doc, pos(1, 0), { indent: { left: 300 } }, ME));
  assert.ok((await body(o)).indexOf(`<w:p><w:pPr><w:ind w:left="300"/><w:jc w:val="center"/>${pWas(8)}</w:pPr>${t('cd')}</w:p>`) > 0, await body(o));
  run(o.history, setParaFormat(o.doc, pos(1, 0), { align: null }, ME));
  assert.equal(revisionsOf(o.doc).length, 2);
  run(o.history, setParaFormat(o.doc, pos(1, 0), { indent: null }, ME));
  assert.deepEqual(listed(o.doc), [['paragraphFormat', 'You', '']]);
  assert.ok((await body(o)).indexOf(`<w:p>${t('cd')}</w:p>`) > 0, await body(o));

  // A paragraph that is all the author's (typed in this session, mark and all) needs no record; its neighbours do.
  const own = await open(p('abcd'));
  run(own.history, insertText(own.doc, pos(0, 4), '\nxy\n', undefined, ME));
  run(own.history, setParaFormat(own.doc, range(pos(0, 0), pos(2, 0)), { align: 'center' }, ME));
  const xml = await body(own);
  assert.equal(revisionIds(xml).filter((id) => xml.indexOf(`<w:pPrChange w:id="${id}"`) >= 0).length, 2, xml);
  assert.ok(/<w:p><w:pPr><w:jc w:val="center"\/><w:rPr><w:ins w:id="\d+"[^>]*\/><\/w:rPr><\/w:pPr><w:ins [^>]*><w:r><w:t>xy<\/w:t>/.test(xml), xml);
  checkWritten(await written(own));
});

// --------------------------------------------------------- resolving, undoing

test('rejecting a new record brings the formatting back, accepting drops the record; all at once does the same', async () => {
  const o = await open(p(r('ab', '<w:i/>') + r('cd')) + p('ef', '<w:jc w:val="right"/>'));
  const before = docPicture(o.doc);
  const start = formatted(o.doc);
  const original = await body(o);
  const make = (): void => {
    run(o.history, setRunFormat(o.doc, range(pos(0, 1), pos(0, 3)), { bold: true }, ME));
    run(o.history, setParaFormat(o.doc, pos(1, 0), { align: 'center' }, ME));
  };
  make();
  assert.deepEqual(listed(o.doc), [
    ['format', 'Me', 'bc'],
    ['paragraphFormat', 'Me', ''],
  ]);
  run(o.history, rejectRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.equal(await body(o), `<w:p>${t('ab', '<w:i/>')}${t('cd')}</w:p><w:p><w:pPr><w:jc w:val="center"/>${pWas(3, '<w:jc w:val="right"/>')}</w:pPr>${t('ef')}</w:p>`);
  run(o.history, acceptRevision(o.doc, revisionsOf(o.doc)[0]));
  assert.equal(await body(o), `<w:p>${t('ab', '<w:i/>')}${t('cd')}</w:p><w:p><w:pPr><w:jc w:val="center"/></w:pPr>${t('ef')}</w:p>`);
  assert.deepEqual(listed(o.doc), []);
  // Back to the start, then all at once.
  while (o.history.canUndo) o.history.undo();
  assert.equal(docPicture(o.doc), before);
  make();
  const all = rejectAll(o.doc);
  assert.equal(all.resolved, 2);
  assert.equal(all.skipped, 0);
  run(o.history, all);
  assert.equal(formatted(o.doc), start);
  assert.equal(await body(o), original.replace(/ xml:space="preserve"/g, ''));
  o.history.undo();
  run(o.history, acceptAll(o.doc));
  assert.deepEqual(listed(o.doc), []);
  assert.equal(await body(o), `<w:p>${t('a', '<w:i/>')}${t('b', '<w:b/><w:bCs/><w:i/>')}${t('c', '<w:b/><w:bCs/>')}${t('d')}</w:p><w:p><w:pPr><w:jc w:val="center"/></w:pPr>${t('ef')}</w:p>`);
});

test('a tracked change of formatting is one step to undo and to redo, and undoing all gives the file back', async () => {
  const o = await open(p('abcd') + p('efgh') + tbl(tr(tc('x') + tc('y')), [1000, 1000]) + p('end'));
  const first = await written(o);
  const before = docPicture(o.doc);
  run(o.history, setRunFormat(o.doc, range(pos(0, 1), pos(1, 2)), { bold: true }, ME));
  run(o.history, toggleList(o.doc, range(pos(0, 1), pos(1, 2)), 'number', ME));
  const after = docPicture(o.doc);
  const second = await written(o);
  o.history.undo();
  o.history.undo();
  assert.equal(o.history.canUndo, false);
  assert.equal(docPicture(o.doc), before);
  assert.equal(await written(o), first);
  o.history.redo();
  o.history.redo();
  assert.equal(docPicture(o.doc), after);
  assert.equal(await written(o), second);
  // The id an undone change had is not given again.
  o.history.undo();
  run(o.history, setParaFormat(o.doc, pos(0, 0), { align: 'center' }, ME));
  checkWritten(await written(o));
});

// ------------------------------------------------------------------ losslessness

test('a document that does not track is formatted as before: no record anywhere', async () => {
  const o = await open(p('abcd') + p('ef'), { settings: '' });
  run(o.history, setRunFormat(o.doc, range(pos(0, 1), pos(0, 3)), { bold: true }));
  run(o.history, setParaFormat(o.doc, pos(1, 0), { align: 'center' }));
  run(o.history, toggleList(o.doc, pos(1, 0), 'bullet'));
  const xml = await written(o);
  assert.equal(xml.indexOf('Change'), -1, xml);
  assert.deepEqual(listed(o.doc), []);
  // Told to track, the command does.
  run(o.history, setRunFormat(o.doc, range(pos(0, 0), pos(0, 1)), { italic: true }, { ...ME, trackChanges: true }));
  assert.deepEqual(listed(o.doc), [['format', 'Me', 'a']]);
  // Undoing everything gives the bytes of the file.
  while (o.history.canUndo) o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes);
});

test('a tracked change of formatting rewrites the paragraphs it touches and nothing else: the records of others stay byte for byte', async () => {
  // Spelled as no writer of this engine would: single quotes, spaces, an attribute the engine does not know.
  const odd =
    "<w:p><w:pPr><w:jc  w:val='center' /><w:pPrChange w:id='31' w:author='You' w:date='2020-01-01T00:00:00Z' w:unknown='x'><w:pPr><w:jc w:val='left'/></w:pPr></w:pPrChange></w:pPr>" +
    "<w:r><w:rPr><w:b/><w:rPrChange w:id='32' w:author='You' w:date='2020-01-01T00:00:00Z'><w:rPr ><w:i/></w:rPr></w:rPrChange></w:rPr><w:t>theirs</w:t></w:r></w:p>";
  const o = await open(p('abcd') + odd + p('efgh') + odd.replace("'31'", "'41'").replace("'32'", "'42'"));
  run(o.history, setRunFormat(o.doc, range(pos(0, 1), pos(0, 3)), { bold: true }, ME));
  run(o.history, setParaFormat(o.doc, pos(2, 0), { align: 'right' }, ME));
  const xml = await written(o);
  assert.ok(xml.indexOf(odd) > 0, xml);
  assert.ok(xml.indexOf(odd.replace("'31'", "'41'").replace("'32'", "'42'")) > 0, xml);
  assert.ok(xml.indexOf(was(43)) > 0 && xml.indexOf(pWas(44)) > 0, xml);
  // A paragraph of theirs that is touched keeps their records as they were read.
  run(o.history, setRunFormat(o.doc, range(pos(1, 0), pos(1, 6)), { underline: true }, ME));
  const touched = await written(o);
  assert.ok(touched.indexOf("<w:pPrChange w:id='31' w:author='You' w:date='2020-01-01T00:00:00Z' w:unknown='x'><w:pPr><w:jc w:val='left'/></w:pPr></w:pPrChange>".replace(/'/g, '"')) > 0, touched);
  assert.ok(touched.indexOf('<w:rPrChange w:id="32" w:author="You" w:date="2020-01-01T00:00:00Z"><w:rPr><w:i/></w:rPr></w:rPrChange>') > 0, touched);
  checkWritten(touched);
});

test('a part that names its elements differently gets records in its own spelling', async () => {
  const o = await open(p('abcd'));
  // The body as a part with another prefix would be read: the model says so in the attributes of its root.
  const source = decoder.decode(await o.pkg.readPart('/word/document.xml'));
  const renamed = source
    .replace(/<(\/?)w:/g, '<$1x:')
    .replace(/ w:/g, ' x:')
    .replace('xmlns:w=', 'xmlns:x=');
  o.pkg.writePart('/word/document.xml', new TextEncoder().encode(renamed));
  const pkg = await OpcPackage.open(await o.pkg.save());
  const doc = await readDocx(pkg);
  const history = new DocHistory<DocRange | undefined>(doc);
  run(history, setRunFormat(doc, range(pos(0, 1), pos(0, 3)), { bold: true }, { ...ME, trackChanges: true }));
  run(history, setParaFormat(doc, pos(0, 0), { align: 'center' }, { ...ME, trackChanges: true }));
  await writeDocx(pkg, doc);
  const xml = decoder.decode(await pkg.readPart('/word/document.xml'));
  assert.ok(xml.indexOf(`<x:rPrChange x:id="1" x:author="Me" x:date="${DATE}"><x:rPr/></x:rPrChange>`) > 0, xml);
  assert.ok(xml.indexOf(`<x:pPrChange x:id="2" x:author="Me" x:date="${DATE}"><x:pPr/></x:pPrChange>`) > 0, xml);
  assert.deepEqual(schemaOrderViolations(xml, 'x:'), []);
  const again = await readDocx(await OpcPackage.open(await pkg.save()));
  assert.deepEqual(listed(again), [
    ['paragraphFormat', 'Me', ''],
    ['format', 'Me', 'bc'],
  ]);
  run(new DocHistory<DocRange | undefined>(again), rejectAll(again));
  assert.deepEqual(listed(again), []);
  assert.equal(picture((paragraphAt(again, 0) as Paragraph).props), '{}');
  assert.equal(picture(((paragraphAt(again, 0) as Paragraph).children[0] as Run).props), '{}');
});

test('the setup of the document names the author and the clock of a record', async () => {
  const o = await open(p('abcd'), { settings: '' });
  setTracking(o.doc, { force: true, author: 'Ada', now: () => new Date(Date.UTC(2026, 0, 2, 3, 4, 5, 678)) });
  run(o.history, setRunFormat(o.doc, range(pos(0, 0), pos(0, 4)), { bold: true }));
  run(o.history, setParaFormat(o.doc, pos(0, 0), { align: 'center' }));
  const who = ' w:author="Ada" w:date="2026-01-02T03:04:05Z"';
  assert.equal(await body(o), `<w:p><w:pPr><w:jc w:val="center"/>${pWas(2, '', who)}</w:pPr>${t('abcd', '<w:b/><w:bCs/>' + was(1, '', who))}</w:p>`);
  const anonymous = await open(p('abcd'));
  setTracking(anonymous.doc, { now: () => new Date(0) });
  run(anonymous.history, setParaFormat(anonymous.doc, pos(0, 0), { align: 'center' }));
  assert.ok((await body(anonymous)).indexOf(pWas(1, '', ' w:author="Author" w:date="1970-01-01T00:00:00Z"')) > 0, await body(anonymous));
});

test('a document read from RTF is formatted with tracked changes too, and the docx made of it holds the records', async () => {
  const doc = readRtf(new TextEncoder().encode('{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\pard\\f0 abcd\\par efgh\\par}'));
  const history = new DocHistory<DocRange | undefined>(doc);
  const before = formatted(doc);
  const track = { ...ME, trackChanges: true };
  run(history, setRunFormat(doc, range(pos(0, 1), pos(0, 3)), { bold: true }, track));
  run(history, setParaFormat(doc, pos(1, 0), { align: 'center' }, track));
  assert.deepEqual(listed(doc), [
    ['format', 'Me', 'bc'],
    ['paragraphFormat', 'Me', ''],
  ]);
  const pkg = await OpcPackage.open(await buildDocx(doc));
  const xml = decoder.decode(await pkg.readPart('/word/document.xml'));
  checkWritten(xml);
  assert.equal(revisionIds(xml).length, 2, xml);
  const again = await readDocx(pkg);
  assert.deepEqual(listed(again), listed(doc));
  run(new DocHistory<DocRange | undefined>(again), rejectAll(again));
  assert.deepEqual(texts(again), ['abcd', 'efgh']);
  assert.equal((await written({ pkg, doc: again, history: new DocHistory<DocRange | undefined>(again), bytes: new Uint8Array() })).indexOf('Change'), -1);
  // In the model that was read from RTF, rejecting gives the formatting back.
  run(history, rejectAll(doc));
  assert.equal(formatted(doc), before);
});

// ------------------------------------------------------------ review findings

const ZED = ' w:author="Zed" w:date="2020-01-01T00:00:00Z"';

test('rejecting a record of the paragraph keeps the tracked changes of others its properties hold', async () => {
  // A list number somebody inserted as a tracked change: the model keeps it as XML inside the properties.
  for (const inner of [`<w:ins w:id="77"${ZED}/>`, `<w:numberingChange w:id="77"${ZED} w:original="%1:1:0:."/>`]) {
    const numPr = `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/>${inner}</w:numPr>`;
    for (const how of ['one', 'all', 'level']) {
      const o = await open(p('abc', numPr) + p('x'));
      assert.deepEqual(listed(o.doc), [['other', undefined, '']], how);
      if (how === 'level') run(o.history, indentList(o.doc, pos(0, 1), 1, ME));
      else run(o.history, setParaFormat(o.doc, pos(0, 1), { align: 'center' }, ME));
      assert.equal(revisionsOf(o.doc).length, 2, how);
      if (how === 'all') {
        const all = rejectAll(o.doc);
        assert.equal(all.skipped, 1, how);
        run(o.history, all);
      } else run(o.history, rejectRevision(o.doc, revisionsOf(o.doc).filter((revision) => revision.author === 'Me')[0]));
      assert.deepEqual(listed(o.doc), [['other', undefined, '']], how);
      assert.equal(await body(o), `<w:p><w:pPr>${numPr}</w:pPr>${t('abc')}</w:p><w:p>${r('x')}</w:p>`, how);
      checkWritten(await written(o), how);
    }
    // A record that says the paragraph was in no list: rejected, the number stays all the same, with the change it carries.
    const listed2 = await open(p('abc', `${numPr}<w:jc w:val="center"/><w:pPrChange w:id="78"${ZED}><w:pPr/></w:pPrChange>`) + p('x'));
    run(listed2.history, rejectRevision(listed2.doc, revisionsOf(listed2.doc).filter((revision) => revision.kind === 'paragraphFormat')[0]));
    assert.equal(await body(listed2), `<w:p><w:pPr>${numPr}</w:pPr>${t('abc')}</w:p><w:p>${r('x')}</w:p>`);
    assert.deepEqual(listed((await reopened(listed2))), [['other', undefined, '']]);
    // Taking the paragraph out of its list would take their change with the number: that cannot be tracked.
    const o = await open(p('abc', numPr));
    const before = docPicture(o.doc);
    refused(() => toggleList(o.doc, pos(0, 1), 'bullet', ME));
    assert.equal(docPicture(o.doc), before);
  }
});

test('rejecting a record of a paragraph mark keeps what the mark says of itself, and what its properties keep', async () => {
  const mark = `<w:rPr><w:ins w:id="5"${ZED}/></w:rPr>`;
  const o = await open(p('ab', mark) + p('cd'));
  run(o.history, setRunFormat(o.doc, range(pos(0, 1), pos(1, 1)), { bold: true }, ME));
  assert.deepEqual(listed(o.doc), [
    ['format', 'Me', 'b'],
    ['insert', 'Zed', '¶'],
    // The formatting of the mark and of the text after it: one change.
    ['format', 'Me', 'c'],
  ]);
  for (let guard = 0; guard < 5; guard++) {
    const mine = revisionsOf(o.doc).filter((revision) => revision.author === 'Me');
    if (!mine.length) break;
    const count = revisionsOf(o.doc).length;
    run(o.history, rejectRevision(o.doc, mine[0]));
    assert.equal(revisionsOf(o.doc).length, count - 1);
  }
  assert.deepEqual(listed(o.doc), [['insert', 'Zed', '¶']]);
  assert.equal(await body(o), `<w:p><w:pPr>${mark}</w:pPr>${t('ab')}</w:p><w:p>${t('cd')}</w:p>`);
});

test('a docx built for a package without a theme states plain colours and fonts inside the records too', async () => {
  const theme =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements><a:clrScheme name="T"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="112233"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="AA0000"/></a:accent1><a:accent2><a:srgbClr val="00BB00"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="T"><a:majorFont><a:latin typeface="Georgia"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Verdana"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="T"/></a:themeElements></a:theme>';
  const themed = '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:color w:val="123456" w:themeColor="accent1"/>';
  const shd = '<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF" w:themeFill="accent2"/>';
  // One record made here, and two that came with the file.
  const theirs = p(r('theirs', `<w:b/><w:rPrChange w:id="60"${ZED}><w:rPr>${themed}</w:rPr></w:rPrChange>`), `<w:jc w:val="right"/><w:pPrChange w:id="61"${ZED}><w:pPr>${shd}</w:pPr></w:pPrChange>`);
  const o = await open(p(r('abcdef', themed), shd) + theirs, { theme });
  run(o.history, setRunFormat(o.doc, range(pos(0, 1), pos(0, 4)), { bold: true }, ME));
  run(o.history, setParaFormat(o.doc, pos(0, 1), { align: 'center' }, ME));
  const pkg = await OpcPackage.open(await buildDocx(o.doc));
  const xml = decoder.decode(await pkg.readPart('/word/document.xml'));
  checkWritten(xml);
  assert.equal(xml.indexOf('heme'), -1, xml);
  const plain = '<w:rPr><w:rFonts w:ascii="Verdana" w:hAnsi="Verdana"/><w:color w:val="AA0000"/></w:rPr>';
  const plainShd = '<w:pPr><w:shd w:val="clear" w:color="auto" w:fill="00BB00"/></w:pPr>';
  assert.ok(xml.indexOf(`<w:rPrChange w:id="62"${WHO}>${plain}</w:rPrChange>`) > 0, xml);
  assert.ok(xml.indexOf(`<w:pPrChange w:id="63"${WHO}>${plainShd}</w:pPrChange>`) > 0, xml);
  assert.ok(xml.indexOf(`<w:rPrChange w:id="60"${ZED}>${plain}</w:rPrChange>`) > 0, xml);
  assert.ok(xml.indexOf(`<w:pPrChange w:id="61"${ZED}>${plainShd}</w:pPrChange>`) > 0, xml);
  // Rejected in the document that was built, the formatting is what it looked like.
  const again = await readDocx(pkg);
  run(new DocHistory<DocRange | undefined>(again), rejectAll(again));
  assert.equal(picture((paragraphAt(again, 0).children[0] as Run).props), '{"color":{"rgb":"AA0000"},"fonts":{"ascii":"Verdana","hAnsi":"Verdana"}}');
  // Written in place, the records of the file stay as they are, theme and all.
  const inPlace = await written(o);
  assert.ok(inPlace.indexOf(`<w:rPrChange w:id="60"${ZED}><w:rPr>${themed}</w:rPr></w:rPrChange>`) > 0, inPlace);
  assert.ok(inPlace.indexOf(`<w:rPrChange w:id="62"${WHO}><w:rPr>${themed}</w:rPr></w:rPrChange>`) > 0, inPlace);
});

test('a paragraph the author split still holds text that was there: formatting it is recorded, mark and all', async () => {
  const o = await open(p('abcdef') + p('next'));
  run(o.history, splitParagraph(o.doc, pos(0, 3), ME));
  run(o.history, setParaFormat(o.doc, pos(0, 1), { align: 'center' }, ME));
  const markIns = `<w:ins w:id="1"${WHO}/>`;
  assert.equal(await body(o), `<w:p><w:pPr><w:jc w:val="center"/><w:rPr>${markIns}</w:rPr>${pWas(2)}</w:pPr>${t('abc')}</w:p><w:p>${t('def')}</w:p><w:p>${r('next')}</w:p>`);
  // The mark is the author's, the text before it is not: bold over both is recorded on both.
  run(o.history, setRunFormat(o.doc, range(pos(0, 2), pos(1, 0)), { bold: true }, ME));
  const b = '<w:b/><w:bCs/>';
  assert.equal(await body(o), `<w:p><w:pPr><w:jc w:val="center"/><w:rPr>${markIns}${b}${was(4)}</w:rPr>${pWas(2)}</w:pPr>${t('ab')}${t('c', b + was(3))}</w:p><w:p>${t('def')}</w:p><w:p>${r('next')}</w:p>`);
  checkWritten(await written(o));
  // The centring can be rejected by itself; the split stays.
  run(o.history, rejectRevision(o.doc, revisionsOf(o.doc).filter((revision) => revision.kind === 'paragraphFormat')[0]));
  assert.ok((await body(o)).indexOf(`<w:p><w:pPr><w:rPr>${markIns}${b}${was(4)}</w:rPr></w:pPr>`) === 0, await body(o));
  run(o.history, rejectAll(o.doc));
  assert.equal(await body(o), `<w:p>${t('abcdef')}</w:p><w:p>${r('next')}</w:p>`);

  // A paragraph that is the author's from its first letter to its mark gets no record at all.
  const own = await open(p('abc') + p('next'));
  run(own.history, insertText(own.doc, pos(0, 3), '\nnew line\n', undefined, ME));
  assert.deepEqual(texts(own.doc), ['abc', 'new line', '', 'next']);
  const count = revisionsOf(own.doc).length;
  run(own.history, setParaFormat(own.doc, pos(1, 2), { align: 'center' }, ME));
  run(own.history, setRunFormat(own.doc, range(pos(1, 0), pos(2, 0)), { bold: true }, ME));
  assert.equal(revisionsOf(own.doc).length, count);
  assert.equal((await written(own)).indexOf('Change'), -1, await written(own));
  // Somebody else formatting it does leave one.
  run(own.history, setParaFormat(own.doc, pos(1, 2), { align: 'right' }, YOU));
  assert.equal(revisionsOf(own.doc).filter((revision) => revision.kind === 'paragraphFormat').length, 1);
});

test('an insertion that came with the file is nobody\'s of this session, whatever name it carries', async () => {
  const mine = (inner: string) => `<w:ins w:id="5"${WHO}>${inner}</w:ins>`;
  const o = await open(p(mine(r('abcd')), `<w:rPr><w:ins w:id="6"${WHO}/></w:rPr>`) + p('x'));
  run(o.history, setRunFormat(o.doc, range(pos(0, 1), pos(0, 3)), { bold: true }, ME));
  run(o.history, setParaFormat(o.doc, pos(0, 1), { align: 'center' }, ME));
  const xml = await body(o);
  assert.ok(xml.indexOf(was(7)) > 0 && xml.indexOf(pWas(8)) > 0, xml);
  checkWritten(await written(o));
});

test('an editor without a name gets a record for text in an insertion of "Author"', async () => {
  const o = await open(p('abcd'));
  setTracking(o.doc, { now: () => new Date(0) });
  run(o.history, insertText(o.doc, pos(0, 2), 'XY'));
  run(o.history, setRunFormat(o.doc, range(pos(0, 2), pos(0, 4)), { bold: true }));
  assert.deepEqual(listed(o.doc), [
    ['insert', 'Author', 'XY'],
    ['format', 'Author', 'XY'],
  ]);
});

test('a record that got another id because its run or its paragraph was cut is still its author\'s to take away', async () => {
  // Three runs with the same properties share the record a command makes; the second and third get ids of their own.
  const all = await open(p(r('ab', '<w:i/>') + r('-') + r('cd', '<w:i/>') + r('-') + r('ef', '<w:i/>')));
  run(all.history, setRunFormat(all.doc, range(pos(0, 0), pos(0, 8)), { bold: true }, ME));
  assert.deepEqual(revisionIds(await body(all)), ['1', '2', '3', '4', '5']);
  run(all.history, setRunFormat(all.doc, range(pos(0, 6), pos(0, 8)), { bold: null }, ME));
  assert.deepEqual(revisionIds(await body(all)), ['1', '2', '3', '4']);
  assert.ok((await body(all)).indexOf(`${r('ef', '<w:i/>').replace(' xml:space="preserve"', '')}</w:p>`) > 0, await body(all));

  // A paragraph with a record of this session, split: both halves hold it, and each half is taken back alone.
  const cut = await open(p('abcdef'));
  run(cut.history, setParaFormat(cut.doc, pos(0, 1), { align: 'center' }, ME));
  run(cut.history, splitParagraph(cut.doc, pos(0, 3), ME));
  assert.deepEqual(revisionIds(await body(cut)), ['3', '2', '1']);
  run(cut.history, setParaFormat(cut.doc, pos(0, 1), { align: null }, ME));
  assert.deepEqual(revisionIds(await body(cut)), ['3', '1']);
  run(cut.history, setParaFormat(cut.doc, pos(1, 1), { align: null }, ME));
  assert.deepEqual(listed(cut.doc), [['insert', 'Me', '¶']]);
  assert.equal(await body(cut), `<w:p><w:pPr><w:rPr><w:ins w:id="3"${WHO}/></w:rPr></w:pPr>${t('abc')}</w:p><w:p>${t('def')}</w:p>`);
});

test('a caret in an empty paragraph: a property turned on and off again leaves nothing behind, tracked or not', async () => {
  for (const settings of [TRACK, '']) {
    const o = await open(p('abc') + p('') + p('styled', '<w:pStyle w:val="Heading1"/>'), { settings });
    const original = await body(o);
    const caret = range(pos(1, 0), pos(1, 0));
    run(o.history, toggleRunFormat(o.doc, caret, 'bold', ME));
    assert.ok((await body(o)).indexOf('<w:b/>') > 0, await body(o));
    assert.equal(revisionsOf(o.doc).length, settings ? 1 : 0);
    run(o.history, toggleRunFormat(o.doc, caret, 'bold', ME));
    assert.deepEqual(listed(o.doc), []);
    const same = (xml: string): string => xml.replace(/ xml:space="preserve"/g, '').replace('<w:p></w:p>', '<w:p/>');
    assert.equal(same(await body(o)), same(original), settings);
    assert.equal((await body(o)).indexOf('w:val="0"'), -1, settings);
  }
});

// --------------------------------------------------------------- still refused

test('what still cannot be tracked is refused, and changes nothing', async () => {
  // (Tables, rows, columns, pictures, links and blocks are tracked: test/doc-track-tables.test.ts, test/doc-track-content.test.ts.)
  const o = await open(p('abcd') + tbl(tr(tc('x', '<w:gridSpan w:val="2"/>')) + tr(tc('z') + tc('w')), [1000, 1000]) + p('end'));
  const before = docPicture(o.doc);
  const cell: DocPosition = { story: 'body', path: [1, 1, 0, 0], offset: 0 };
  // A column through a merged cell changes that cell; a block that is neither a paragraph nor a table has no mark to say it was inserted.
  refused(() => insertColumn(o.doc, cell, 'right'));
  refused(() => deleteColumn(o.doc, cell));
  refused(() => insertBlock(o.doc, pos(0, 1), { type: 'mark', mark: { type: 'bookmarkStart', id: '9', name: 'b' } }));
  assert.equal(docPicture(o.doc), before);
  assert.equal(o.doc.body.modified, undefined);
  // Formatting inside a cell is tracked like anywhere else.
  run(o.history, setRunFormat(o.doc, range(cell, { ...cell, offset: 1 }), { bold: true }, ME));
  run(o.history, setParaFormat(o.doc, cell, { align: 'center' }, ME));
  assert.deepEqual(listed(o.doc), [
    ['paragraphFormat', 'Me', ''],
    ['format', 'Me', 'z'],
  ]);
  checkWritten(await written(o));
});

// ------------------------------------------------------------------------ random

/** The text and the formatting of every paragraph of a document, run boundaries aside: what rejecting everything has to bring back. */
function formatted(doc: DocModel, withoutKept = false): string {
  const out: unknown[] = [];
  for (const story of storyIds(doc)) {
    for (const path of paragraphPaths(doc, story)) {
      const paragraph = paragraphAt(doc, path, story);
      const spans: [string, string][] = [];
      for (const inline of paragraph.children) {
        if (inline.type !== 'run') continue;
        let text = '';
        for (const item of inline.content) text += item.type === 'text' ? item.text : `<${item.type}>`;
        if (text === '') continue;
        const props = picture(inline.props);
        const last = spans[spans.length - 1];
        if (last && last[1] === props) last[0] += text;
        else spans.push([text, props]);
      }
      out.push([story, path.join('.'), picture(withoutKept ? { ...paragraph.props, extras: undefined } : paragraph.props), spans]);
    }
  }
  return JSON.stringify(out);
}

const RUN_PATCHES = [{ bold: true }, { bold: null }, { italic: true }, { italic: null }, { underline: true }, { size: 30 }, { size: null }, { color: 'FF0000' }, { highlight: 'yellow' }, { style: 'Strong' }, { style: null }, { font: 'Other' }];
const PARA_PATCHES = [{ align: 'center' }, { align: null }, { indent: { left: 400 } }, { indent: null }, { spacing: { before: 100 } }, { style: 'Quote' }, { style: null }, { keepNext: true }];

/** The kind of the change the last randomTracked() accepted or rejected alone; undefined when it did something else. */
let resolvedKind: string | undefined;

/** The ids of the tracked changes in the XML that are no records of changed formatting, in order. */
function otherIds(xml: string): string[] {
  const out: string[] = [];
  for (const match of xml.matchAll(/<w:(?:ins|del|moveFrom|moveTo|numberingChange)\b[^>]*? w:id="(\d+)"/g)) out.push(match[1]);
  return out.sort();
}

/** How many changes the document lists, and how many of them are changes the model only keeps. */
function counted(doc: DocModel): { all: number; kept: number } {
  let kept = 0;
  const list = revisionsOf(doc);
  for (const revision of list) if (revision.kind === 'other') kept += revision.count;
  return { all: list.length, kept };
}

function randomTracked(o: Opened, d: ReturnType<typeof dice>, resolve: boolean): DocEdit | 'undo' | undefined {
  resolvedKind = undefined;
  const story = d.pick(storyIds(o.doc));
  const paths = paragraphPaths(o.doc, story);
  const at = randomPosition(o.doc, d, story, paths);
  const other = randomPosition(o.doc, d, story, paths);
  const who: DocEditOptions = resolve && d.chance(0.25) ? YOU : ME;
  const kind = d.int(resolve ? 14 : 12);
  switch (kind) {
    case 0:
      return insertText(o.doc, at, d.pick(['x', 'two words', 'a\nb', '\t', ' ']), undefined, who);
    case 1:
      return deleteBackward(o.doc, at, 'char', who);
    case 2:
      return deleteForward(o.doc, at, d.chance(0.3) ? 'word' : 'char', who);
    case 3:
      return deleteRange(o.doc, range(at, other), who);
    case 4:
      return splitParagraph(o.doc, at, who);
    case 5:
    case 6:
    case 7:
      return setRunFormat(o.doc, range(at, other), d.pick(RUN_PATCHES), who);
    case 8:
      return toggleRunFormat(o.doc, range(at, other), d.pick(['bold', 'italic', 'underline', 'strike'] as const), who);
    case 9:
      return setParaFormat(o.doc, d.chance(0.5) ? at : range(at, other), d.pick(PARA_PATCHES), who);
    case 10:
      return d.chance(0.5) ? toggleList(o.doc, range(at, other), d.pick(['bullet', 'number'] as const), who) : indentList(o.doc, at, d.chance(0.5) ? 1 : -1, who);
    case 11:
      return 'undo';
    default: {
      const list = listRevisions(o.doc).list.filter((revision) => (kind === 12 ? revision.canAccept : revision.canReject));
      if (!list.length) return undefined;
      const revision = d.pick(list);
      resolvedKind = revision.kind;
      return kind === 12 ? acceptRevision(o.doc, revision) : rejectRevision(o.doc, revision);
    }
  }
}

const RANDOM_BODY =
  p('alpha beta gamma') +
  p(r('bold', '<w:b/>') + r(' and plain') + r(' italic', '<w:i/>'), '<w:jc w:val="right"/>') +
  p('') +
  p('item one', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
  tbl(tr(tc('one') + tc(p('two') + p(r('three', '<w:u w:val="single"/>')))), [1500, 1500]) +
  p('omega end', '<w:pStyle w:val="Heading1"/><w:spacing w:before="120"/>');

test('random tracked typing, deleting and formatting by one author: every file is sound, and rejecting everything gives the document back as it was', async () => {
  for (let seed = 1; seed <= 16; seed++) {
    const d = dice(seed * 6271);
    const o = await open(RANDOM_BODY, { headers: { rIdH: p('header text') } });
    const before = formatted(o.doc);
    const numbering = picture(o.doc.numbering);
    let made = 0;
    for (let step = 0; step < 45; step++) {
      let edit: DocEdit | 'undo' | undefined;
      try {
        edit = randomTracked(o, d, false);
      } catch (error) {
        // A range over a table, a place that cannot be edited: nothing happened.
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
      if (step % 9 === 0) {
        checkWritten(await written(o), `seed ${seed} step ${step}`);
        for (const part of ['/word/header1.xml']) checkWritten(await written(o, part), `seed ${seed} step ${step} ${part}`);
      }
    }
    assert.ok(made > 10, `seed ${seed}: ${made} edits`);
    const xml = await written(o);
    checkWritten(xml, `seed ${seed}`);
    // What is written is what is read again.
    const again = await reopened(o);
    assert.equal(formatted(again), formatted(o.doc), `seed ${seed}: read again`);
    assert.deepEqual(listed(again), listed(o.doc), `seed ${seed}: revisions read again`);
    // Everything rejected: the text and the formatting of the start. (Lists made on the way stay defined, unused.)
    const all = rejectAll(o.doc);
    assert.equal(all.skipped, 0, `seed ${seed}`);
    run(o.history, all);
    assert.deepEqual(listed(o.doc), [], `seed ${seed}`);
    assert.equal(formatted(o.doc), before, `seed ${seed}`);
    checkWritten(await written(o), `seed ${seed} rejected`);
    assert.equal((await written(o)).indexOf('Change'), -1, `seed ${seed}`);
    // And undoing all of it gives the file back.
    while (o.history.canUndo) o.history.undo();
    assert.equal(picture(o.doc.numbering), numbering, `seed ${seed}`);
    await writeDocx(o.pkg, o.doc);
    assert.deepEqual(await o.pkg.save(), o.bytes, `seed ${seed}: bytes`);
  }
});

test('random tracked edits by two authors with single verdicts in between: every file is sound, and accepting everything leaves no change', async () => {
  for (let seed = 1; seed <= 16; seed++) {
    const d = dice(seed * 9173);
    const theirs = p(r('theirs', '<w:b/><w:rPrChange w:id="50" w:author="Zed" w:date="2020-01-01T00:00:00Z"><w:rPr><w:i/></w:rPr></w:rPrChange>'), '<w:jc w:val="center"/><w:pPrChange w:id="51" w:author="Zed" w:date="2020-01-01T00:00:00Z"><w:pPr/></w:pPrChange>');
    // A list number somebody inserted as a tracked change: kept as XML in the properties of its paragraph.
    const numbered = p('numbered', `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/><w:ins w:id="52" w:author="Zed" w:date="2020-01-01T00:00:00Z"/></w:numPr>`);
    const o = await open(RANDOM_BODY + theirs + numbered);
    let made = 0;
    for (let step = 0; step < 50; step++) {
      let edit: DocEdit | 'undo' | undefined;
      try {
        edit = randomTracked(o, d, true);
      } catch (error) {
        assert.ok(error instanceof RangeError, `seed ${seed} step ${step}: ${String(error)}`);
        continue;
      }
      if (edit === 'undo') {
        if (o.history.canUndo && d.chance(0.5)) o.history.undo();
        continue;
      }
      if (!edit || edit.ops.length === 0) continue;
      const alone = resolvedKind;
      const before = counted(o.doc);
      // What the file holds of changes that are no records of formatting, before a record is resolved alone.
      const others = alone === 'format' || alone === 'paragraphFormat' ? otherIds(await written(o)) : undefined;
      run(o.history, edit);
      made++;
      checkInvariants(o.doc, false);
      const after = counted(o.doc);
      // A change of formatting that is accepted or rejected alone takes nothing with it: every insertion, deletion
      // and change the model only keeps (the inserted list number) is still in the file.
      if (others) {
        // (The list may get shorter by more than one: two deletions with nothing but the record between them are one change now.)
        assert.ok(after.all < before.all, `seed ${seed} step ${step}: ${alone} resolved alone, ${before.all} changes before and ${after.all} after`);
        assert.deepEqual(otherIds(await written(o)), others, `seed ${seed} step ${step}: ${alone} resolved alone`);
      }
      if (step % 7 === 0) checkWritten(await written(o), `seed ${seed} step ${step}`);
    }
    assert.ok(made > 10, `seed ${seed}: ${made} edits`);
    checkWritten(await written(o), `seed ${seed}`);
    const again = await reopened(o);
    // (The list number the model keeps as XML is kept as it was read, under the fields that say what it is now: left out.)
    assert.equal(formatted(again, true), formatted(o.doc, true), `seed ${seed}: read again`);
    assert.deepEqual(listed(again), listed(o.doc), `seed ${seed}: revisions read again`);
    const kept = counted(o.doc).kept;
    const all = acceptAll(o.doc);
    assert.equal(all.skipped, kept, `seed ${seed}`);
    run(o.history, all);
    // What the model only keeps is passed over, and still there; nothing else is.
    assert.deepEqual(listed(o.doc).filter((revision) => revision[0] !== 'other'), [], `seed ${seed}`);
    // (Fewer of them than before where an accepted deletion of a paragraph mark joined two paragraphs that both held one.)
    const left = counted(o.doc).kept;
    assert.ok(left >= 1 && left <= kept, `seed ${seed}: ${left} kept changes of ${kept}`);
    const xml = await written(o);
    checkWritten(xml, `seed ${seed} accepted`);
    assert.equal(revisionIds(xml).length, left, `seed ${seed}: ${revisionIds(xml).join(' ')}`);
  }
});
