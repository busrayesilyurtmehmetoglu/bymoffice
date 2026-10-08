// Third review of tracked changes: a tracked deletion in the author's own insertion takes out
// nothing but text that sits directly in it, and joins two paragraphs for good only when the
// paragraph that would go is formatted like the one that stays.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import { type DocModel } from '../src/model/doc/types';
import { DocHistory, deleteRange, insertText, type DocEditOptions, type DocRange } from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { parseXml } from '../src/xml';
import { NUMBERING, STYLES, docPicture, p, paragraphAt, pos, r, range, run, texts } from './doc-ops-fixtures';
import { buildDocx as buildSpec, type DocxSpec } from './docx-fixtures';

const decoder = new TextDecoder();
const ME: DocEditOptions = { author: 'Me', date: '2026-10-05T10:00:00Z' };
const TRACK = '<w:trackRevisions/>';
const SECT = '<w:sectPr><w:pgSz w:w="9000" w:h="9000"/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr>';
const OWN = (inner: string, id = 7, author = 'Me') => `<w:ins w:id="${id}" w:author="${author}" w:date="2020-01-01T00:00:00Z">${inner}</w:ins>`;

interface Opened {
  pkg: OpcPackage;
  doc: DocModel;
  history: DocHistory<DocRange | undefined>;
}

async function open(body: string, extra: Partial<DocxSpec> = {}): Promise<Opened> {
  const bytes = await buildSpec({ body: body + SECT, styles: STYLES, settings: TRACK, ...extra });
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  return { pkg, doc, history: new DocHistory<DocRange | undefined>(doc) };
}

async function written(o: Opened): Promise<string> {
  await writeDocx(o.pkg, o.doc);
  return decoder.decode(await o.pkg.readPart('/word/document.xml'));
}

async function body(o: Opened): Promise<string> {
  const text = await written(o);
  return text.slice(text.indexOf('<w:body>') + 8, text.indexOf('<w:sectPr>'));
}

const count = (text: string, part: string): number => text.split(part).length - 1;

const WRAPPED =
  '<w:fldSimple w:instr=" DOCPROPERTY Title ">' +
  r('field') +
  '</w:fldSimple>' +
  '<w:hyperlink r:id="rIdL">' +
  r('link') +
  '</w:hyperlink>' +
  '<w:sdt><w:sdtPr><w:alias w:val="Control"/><w:id w:val="4242"/></w:sdtPr><w:sdtContent>' +
  r('control') +
  '</w:sdtContent></w:sdt>' +
  '<w:smartTag w:uri="urn:example" w:element="tag">' +
  r('tagged') +
  '</w:smartTag>';
const LINK: Partial<DocxSpec> = { rels: [['rIdL', 'hyperlink', 'https://example.org/', true]] };

test('review 3: deleting the whole of an own insertion keeps the field, the link, the control and the tag in it', async () => {
  const o = await open(p(OWN(WRAPPED)), LINK);
  const original = await written(o);
  const picture = docPicture(o.doc);
  const length = texts(o.doc)[0].length;
  assert.equal(texts(o.doc)[0], 'fieldlinkcontroltagged');
  run(o.history, deleteRange(o.doc, range(pos(0, 0), pos(0, length)), ME));
  assert.deepEqual(texts(o.doc), ['']);
  const xml = await body(o);
  for (const kept of ['<w:fldSimple w:instr=" DOCPROPERTY Title ">', '<w:hyperlink r:id="rIdL">', '<w:alias w:val="Control"/>', '<w:id w:val="4242"/>', '<w:smartTag w:uri="urn:example" w:element="tag">']) {
    assert.equal(count(xml, kept), 1, `${kept}\n${xml}`);
  }
  // The text of each is deleted where it was: a deletion inside the wrapper, inside the insertion.
  for (const text of ['field', 'link', 'control', 'tagged']) assert.equal(count(xml, `>${text}</w:delText>`), 1, `${text}\n${xml}`);
  assert.ok(/<w:fldSimple [^>]*><w:del [^>]*><w:r><w:delText[^>]*>field<\/w:delText><\/w:r><\/w:del><\/w:fldSimple>/.test(xml), xml);
  assert.ok(/^<w:p><w:ins w:id="7"[^>]*>.*<\/w:ins><\/w:p>$/.test(xml), xml);
  assert.equal(count(xml, '<w:t'), 0, xml);
  parseXml(await written(o));
  const ids = Array.from(xml.matchAll(/<w:(?:ins|del)\b[^>]*? w:id="(\d+)"/g), (match) => match[1]);
  assert.equal(new Set(ids).size, ids.length, xml);
  o.history.undo();
  assert.equal(docPicture(o.doc), picture);
  assert.equal(await written(o), original);
});

test('review 3: text directly in the own insertion still goes for good, beside wrappers that stay', async () => {
  const o = await open(p(r('a') + OWN(r('typed') + `<w:hyperlink r:id="rIdL">${r('link')}</w:hyperlink>` + r('more')) + r('b')), LINK);
  run(o.history, deleteRange(o.doc, range(pos(0, 1), pos(0, 14)), ME));
  assert.deepEqual(texts(o.doc), ['ab']);
  const xml = await body(o);
  assert.equal(xml.indexOf('typed'), -1, xml);
  assert.equal(xml.indexOf('more'), -1, xml);
  assert.ok(/<w:ins w:id="7"[^>]*><w:hyperlink r:id="rIdL"><w:del [^>]*><w:r><w:delText>link<\/w:delText><\/w:r><\/w:del><\/w:hyperlink><\/w:ins>/.test(xml), xml);
});

test('review 3: an own insertion inside a wrapper goes for good while the wrapper holds something else, and stays deleted when it is all the wrapper holds', async () => {
  // Typed into a link that has other text: taking it back leaves the link as it was.
  const shared = await open(p(`<w:hyperlink r:id="rIdL">${r('the ')}${OWN(r('X'))}${r('site')}</w:hyperlink>`), LINK);
  run(shared.history, deleteRange(shared.doc, range(pos(0, 4), pos(0, 5)), ME));
  assert.equal(await body(shared), `<w:p><w:hyperlink r:id="rIdL"><w:r><w:t>the site</w:t></w:r></w:hyperlink></w:p>`);

  // The insertion is everything the control holds: the control must not go with it.
  const alone = await open(p(r('a') + `<w:sdt><w:sdtPr><w:id w:val="99"/></w:sdtPr><w:sdtContent>${OWN(r('only'))}</w:sdtContent></w:sdt>` + r('b')));
  const picture = docPicture(alone.doc);
  run(alone.history, deleteRange(alone.doc, range(pos(0, 1), pos(0, 5)), ME));
  assert.deepEqual(texts(alone.doc), ['ab']);
  const xml = await body(alone);
  assert.equal(count(xml, '<w:id w:val="99"/>'), 1, xml);
  assert.equal(count(xml, '<w:sdt>'), 1, xml);
  assert.equal(count(xml, '</w:delText>'), 1, xml);
  alone.history.undo();
  assert.equal(docPicture(alone.doc), picture);
});

// ------------------------------------------------------------ paragraph marks

const insertedMark = (pPr: string) => `${pPr}<w:rPr><w:ins w:id="20" w:author="Me" w:date="2020-01-01T00:00:00Z"/></w:rPr>`;

test('review 3: a tracked cut over an own paragraph mark keeps both paragraphs when they are formatted differently', async () => {
  const o = await open(
    p(OWN(r('Hello'), 7), insertedMark('<w:pStyle w:val="Heading1"/>')) + p(OWN(r('World'), 8), '<w:pStyle w:val="Quote"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr><w:jc w:val="center"/>') + p('end'),
    { numbering: NUMBERING },
  );
  const picture = docPicture(o.doc);
  const second = paragraphAt(o.doc, 1).props;
  assert.equal(second.style, 'Quote');
  const at = run(o.history, deleteRange(o.doc, range(pos(0, 3), pos(1, 2)), ME));
  // The paragraphs are still two, and the second is what it was.
  assert.deepEqual(texts(o.doc), ['Hel', 'rld', 'end']);
  const first = paragraphAt(o.doc, 0).props;
  assert.equal(first.style, 'Heading1');
  assert.ok(first.markProps && first.markProps.deleted && first.markProps.inserted, 'the mark is deleted inside its insertion');
  const kept = paragraphAt(o.doc, 1).props;
  assert.equal(kept.style, 'Quote');
  assert.equal(kept.align, second.align);
  assert.deepEqual(kept.numbering, second.numbering);
  // Pasting the lines back where they were cut gives the two paragraphs their own formatting.
  run(o.history, insertText(o.doc, at.focus, 'lo\nWo', undefined, ME));
  const styles = [0, 1, 2].map((index) => paragraphAt(o.doc, index).props.style);
  assert.equal(styles[0], 'Heading1');
  assert.equal(paragraphAt(o.doc, texts(o.doc).length - 2).props.style, 'Quote');
  parseXml(await written(o));
  while (o.history.canUndo) o.history.undo();
  assert.equal(docPicture(o.doc), picture);
});

test('review 3: an own paragraph mark between two paragraphs formatted alike still goes for good', async () => {
  const pPr = '<w:pStyle w:val="Heading1"/><w:jc w:val="center"/>';
  const o = await open(p(OWN(r('Hello'), 7), insertedMark(pPr)) + p(OWN(r('World'), 8), pPr) + p('end'));
  run(o.history, deleteRange(o.doc, range(pos(0, 3), pos(1, 2)), ME));
  assert.deepEqual(texts(o.doc), ['Helrld', 'end']);
  assert.equal(paragraphAt(o.doc, 0).props.style, 'Heading1');
  assert.equal(paragraphAt(o.doc, 0).props.markProps, undefined);
});
