// Find and Replace beyond the body, after its second review: whether alternate content can take a
// change is worked out once for all its graphics (not once per graphic: many boxes in one element
// made that quadratic); the replacement does not take another author's tracked formatting or go
// into another author's tracked insertion, and leaves no empty link or revision behind; the other
// branch of a box Word wrote counts as holding the same text when it differs only in ids, rsids,
// the order of attributes or the hints of the last layout.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import { paragraphText, type DocModel, type Paragraph } from '../src/model/doc/types';
import { OpcPackage } from '../src/opc';
import { DocHistory, docParagraphs, findAllInDocument, listRevisions, rejectAll, replaceMatches, replaceText, type DocEdit } from '../src/ops/doc/index';
import { SECT, STYLES, p, r } from './doc-ops-fixtures';
import { buildDocx } from './docx-fixtures';

const decoder = new TextDecoder();

const ANCHOR_HEAD =
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="1828800" cy="457200"/><wp:wrapNone/>';
const drawing = (inner: string, id = 1) =>
  `<w:drawing>${ANCHOR_HEAD}<wp:docPr id="${id}" name="Box ${id}"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`;
const pict = (inner: string) => `<w:pict><v:shape style="width:144pt;height:36pt"><v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:shape></w:pict>`;
const alt = (choice: string, fallback: string) => `<w:r><mc:AlternateContent><mc:Choice Requires="wps">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent></w:r>`;

interface Opened {
  pkg: OpcPackage;
  doc: DocModel;
}

async function openBody(xml: string): Promise<Opened> {
  const pkg = await OpcPackage.open(await buildDocx({ body: xml + SECT, styles: STYLES }));
  return { pkg, doc: await readDocx(pkg) };
}

async function savedXml(opened: Opened): Promise<{ xml: string; doc: DocModel }> {
  await writeDocx(opened.pkg, opened.doc);
  const pkg = await OpcPackage.open(await opened.pkg.save());
  return { xml: decoder.decode(await pkg.readPart('/word/document.xml')), doc: await readDocx(pkg) };
}

function replaceAll(doc: DocModel, text: string, replacement: string) {
  const history = new DocHistory(doc);
  const result = replaceMatches(doc, findAllInDocument(doc, text), replacement, { apply: (edit: DocEdit) => history.applyEdit({ ...edit, mergeKey: 'all' }, undefined) });
  history.seal();
  return result;
}

const body = (xml: string): string => xml.slice(xml.indexOf('<w:body>') + 8, xml.indexOf('<w:sectPr'));
const count = (text: string, part: string): number => text.split(part).length - 1;

// ------------------------------------------------------------------- speed

test('2,000 text boxes in one alternate content: listed, replaced and saved in time that follows their number', async () => {
  const K = 2000;
  const choice = Array.from({ length: K }, (_, i) => drawing(p('cat ' + i), i + 1)).join('');
  const same = Array.from({ length: K }, (_, i) => pict(p('cat ' + i))).join('');
  // Branches that take the change, and a Fallback without boxes: every box refused.
  for (const [what, fallback, replaced] of [
    ['the same boxes', same, K],
    ['no boxes', '<w:t>x</w:t>', 0],
  ] as [string, string, number][]) {
    const opened = await openBody(p(r('a ') + alt(choice, fallback)));
    const started = Date.now();
    const list = docParagraphs(opened.doc);
    assert.equal(list.length, K + 1, what);
    const result = replaceAll(opened.doc, 'cat', 'dog');
    const saved = await savedXml(opened);
    const took = Date.now() - started;
    assert.equal(result.replaced, replaced, what);
    assert.equal(count(saved.xml, 'dog'), replaced * 2, what);
    // A guard against a blow-up, not a measure of speed (the review measured 22 s here).
    assert.ok(took < 8000, `${what}: ${took} ms`);
  }
});

// ------------------------------------------------- whose the replacement is

test('the replacement does not take the tracked formatting another author left on the matched text', async () => {
  const opened = await openBody(`<w:p>${r('ab')}<w:r><w:rPr><w:b/><w:rPrChange w:id="3" w:author="Other" w:date="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>cd</w:t></w:r></w:p>`);
  assert.equal(replaceAll(opened.doc, 'cd', 'Z').replaced, 1);
  assert.equal(listRevisions(opened.doc).list.length, 0);
  const saved = await savedXml(opened);
  assert.doesNotMatch(saved.xml, /rPrChange/);
  assert.equal(paragraphText(saved.doc.body.blocks[0] as Paragraph), 'abZ');
});

test("the replacement is not written into another author's tracked insertion: rejecting it all leaves the user's text", async () => {
  const opened = await openBody(`<w:p>${r('ab')}<w:ins w:id="7" w:author="Other" w:date="2026-01-01T00:00:00Z">${r('cd')}</w:ins>${r('ef')}</w:p>`);
  assert.equal(replaceAll(opened.doc, 'cd', 'Z').replaced, 1);
  const history = new DocHistory(opened.doc);
  history.applyEdit(rejectAll(opened.doc) as unknown as DocEdit, undefined);
  assert.match(paragraphText(opened.doc.body.blocks[0] as Paragraph), /Z/);
});

test('replaced by nothing: no empty link, revision or wrapper is left; an empty content control stays', async () => {
  const cases: [string, string, RegExp | undefined][] = [
    ['link', `<w:p>${r('ab')}<w:hyperlink w:anchor="x">${r('cd')}</w:hyperlink>${r('ef')}</w:p>`, /<w:hyperlink/],
    ['insertion', `<w:p>${r('ab')}<w:ins w:id="7" w:author="Other" w:date="2026-01-01T00:00:00Z">${r('cd')}</w:ins>${r('ef')}</w:p>`, /<w:ins /],
    ['smart tag', `<w:p>${r('ab')}<w:smartTag w:uri="u" w:element="e">${r('cd')}</w:smartTag>${r('ef')}</w:p>`, /<w:smartTag/],
  ];
  for (const [what, xml, gone] of cases) {
    const opened = await openBody(xml);
    assert.equal(replaceAll(opened.doc, 'cd', '').replaced, 1, what);
    const saved = await savedXml(opened);
    assert.doesNotMatch(body(saved.xml), gone!, what);
    assert.doesNotMatch(body(saved.xml), /<w:t\/>/, what);
    assert.equal(paragraphText(saved.doc.body.blocks[0] as Paragraph), 'abef', what);
  }
  const control = await openBody(`<w:p>${r('ab')}<w:sdt><w:sdtPr><w:tag w:val="T"/></w:sdtPr><w:sdtContent><w:hyperlink w:anchor="x">${r('cd')}</w:hyperlink></w:sdtContent></w:sdt>${r('ef')}</w:p>`);
  assert.equal(replaceAll(control.doc, 'cd', '').replaced, 1);
  const kept = body((await savedXml(control)).xml);
  assert.equal(count(kept, '<w:sdt>'), 1);
  assert.doesNotMatch(kept, /<w:hyperlink/);
});

test('replaceText still writes in place where nothing of the kind is in the way', async () => {
  const opened = await openBody(`<w:p><w:bookmarkStart w:id="1" w:name="m"/>${r('ab')}<w:bookmarkEnd w:id="1"/>${r('cd')}</w:p>`);
  const match = findAllInDocument(opened.doc, 'ab')[0];
  new DocHistory(opened.doc).applyEdit(replaceText(opened.doc, match.range, 'XY'), undefined);
  assert.match(body((await savedXml(opened)).xml), /<w:bookmarkStart w:id="1" w:name="m"\/><w:r><w:t>XY<\/w:t><\/w:r><w:bookmarkEnd w:id="1"\/>/);
});

// ------------------------------------------- the same text in the other branch

test('a Fallback that differs only in paragraph ids, rsids, the order of attributes or layout hints holds the same text', async () => {
  const P = (attrs: string, text: string, run = '') => `<w:p ${attrs}><w:r${run}><w:t>${text}</w:t></w:r></w:p>`;
  const cases: [string, string, string][] = [
    ['paragraph ids', P('w14:paraId="1A" w14:textId="2B"', 'cat'), P('w14:paraId="3C" w14:textId="4D"', 'cat')],
    ['rsid of a paragraph', P('w:rsidR="00AA"', 'cat'), P('w:rsidR="00BB"', 'cat')],
    ['rsid of a run', P('', 'cat', ' w:rsidR="0011"'), P('', 'cat', ' w:rsidR="0022"')],
    ['order of attributes', '<w:p w14:paraId="1" w14:textId="2"><w:r><w:t>cat</w:t></w:r></w:p>', '<w:p w14:textId="2" w14:paraId="1"><w:r><w:t>cat</w:t></w:r></w:p>'],
    ['a hint of the last layout', '<w:p><w:r><w:lastRenderedPageBreak/><w:t>cat</w:t></w:r></w:p>', '<w:p><w:r><w:t>cat</w:t></w:r></w:p>'],
  ];
  for (const [what, choice, fallback] of cases) {
    const opened = await openBody(p(r('a ') + alt(drawing(choice), pict(fallback))));
    const matches = findAllInDocument(opened.doc, 'cat');
    assert.equal(matches[0].replaceable, true, what);
    assert.equal(replaceAll(opened.doc, 'cat', 'dog').replaced, 1, what);
    const saved = await savedXml(opened);
    assert.equal(count(saved.xml, 'cat'), 0, what);
    assert.equal(count(saved.xml, 'dog'), 2, what);
    assert.equal(count(saved.xml, '<mc:AlternateContent>'), 1, what);
  }
  // Other text is still other text.
  const other = await openBody(p(r('a ') + alt(drawing(P('', 'cat')), pict(P('', 'cat!')))));
  assert.equal(findAllInDocument(other.doc, 'cat')[0].replaceable, false);
});
