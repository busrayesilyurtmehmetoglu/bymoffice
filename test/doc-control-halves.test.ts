// Content controls that lie side by side, and the halves of one that Enter cut: two controls stay
// two whatever is edited around them; a control that Enter cut and Backspace joined again is the
// control that was read (a test of the commands and the writer: no room).
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { emptyWriteReport, writeDocx } from '../src/formats/docx/write';
import { OpcPackage } from '../src/opc';
import { DocHistory, acceptAll, deleteBackward, deleteRange, insertText, rejectAll, splitParagraph, type DocEdit } from '../src/ops/doc/index';
import type { DocModel } from '../src/model/doc/types';
import { buildDocx, type DocxSpec } from './docx-fixtures';

const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);
const count = (text: string, pattern: RegExp): number => (text.match(pattern) || []).length;
const A = '<w:sdt><w:sdtPr><w:tag w:val="f"/><w:id w:val="5"/></w:sdtPr><w:sdtContent><w:r><w:t>AAA</w:t></w:r></w:sdtContent></w:sdt>';
const B = '<w:sdt><w:sdtPr><w:tag w:val="f"/></w:sdtPr><w:sdtContent><w:r><w:t>BBB</w:t></w:r></w:sdtContent></w:sdt>';
const tail = '<w:r><w:t>tail</w:t></w:r>';
const p = (inner: string): string => `<w:p>${inner}</w:p>`;
const at = (story: string, path: number[], offset: number) => ({ story, path, offset });

/** Opens a file, makes the edits one after the other, writes it: the part as it was read and as it is written. */
async function edited(spec: Partial<DocxSpec>, part: string, edits: ((doc: DocModel) => DocEdit)[]): Promise<{ before: string; after: string }> {
  const pkg = await OpcPackage.open(await buildDocx(spec));
  const before = decode(await pkg.readPart(part));
  const doc = await readDocx(pkg);
  const history = new DocHistory<undefined>(doc);
  for (const edit of edits) history.applyEdit(edit(doc), undefined);
  await writeDocx(pkg, doc, { report: emptyWriteReport() });
  return { before, after: decode(await (await OpcPackage.open(await pkg.save())).readPart(part)) };
}

test('two controls side by side in a paragraph stay two, whatever is typed beside them or in them, and wherever they stand', async () => {
  const cell = `<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${p(A + B + tail)}</w:tc></w:tr></w:tbl><w:p/>`;
  const cases: [string, Partial<DocxSpec>, string, (doc: DocModel) => DocEdit][] = [
    ['typed behind them', { body: p(A + B + tail) }, '/word/document.xml', (doc) => insertText(doc, at('body', [0], 8), 'X')],
    ['typed in the second', { body: p(A + B + tail) }, '/word/document.xml', (doc) => insertText(doc, at('body', [0], 4), 'X')],
    ['typed in another paragraph', { body: p(A + B + tail) + p('<w:r><w:t>other</w:t></w:r>') }, '/word/document.xml', (doc) => insertText(doc, at('body', [1], 2), 'X')],
    ['in a table cell', { body: cell }, '/word/document.xml', (doc) => insertText(doc, at('body', [0, 0, 0, 0], 8), 'X')],
    ['in a header', { body: p('x') + '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>', headers: { rIdH: p(A + B + tail) } }, '/word/header1.xml', (doc) => insertText(doc, at('header:rIdH', [0], 8), 'X')],
  ];
  for (const [what, spec, part, edit] of cases) {
    const { before, after } = await edited(spec, part, [edit]);
    assert.equal(count(before, /<w:sdt>/g), 2, what);
    assert.equal(count(after, /<w:sdt>/g), 2, `${what}: ${after.slice(after.indexOf('<w:p'), after.indexOf('<w:p') + 600)}`);
    assert.equal(count(after, /<w:id w:val="5"\/>/g), 1, what);
  }
});

test('Enter in a control and Backspace at once: the file is the file that was read, byte for byte', async () => {
  const { before, after } = await edited({ body: p(A + tail) }, '/word/document.xml', [(doc) => splitParagraph(doc, at('body', [0], 1)), (doc) => deleteBackward(doc, at('body', [1], 0))]);
  assert.equal(after, before);
  // Twice Enter in it, and twice Backspace: the same.
  const twice = await edited({ body: p(A + tail) }, '/word/document.xml', [
    (doc) => splitParagraph(doc, at('body', [0], 1)),
    (doc) => splitParagraph(doc, at('body', [1], 1)),
    (doc) => deleteBackward(doc, at('body', [2], 0)),
    (doc) => deleteBackward(doc, at('body', [1], 0)),
  ]);
  assert.equal(twice.after, twice.before);
  // And Enter alone: two controls, the id with the first.
  const split = await edited({ body: p(A + tail) }, '/word/document.xml', [(doc) => splitParagraph(doc, at('body', [0], 1))]);
  assert.equal(count(split.after, /<w:sdt>/g), 2);
  assert.equal(count(split.after, /<w:id w:val="5"\/>/g), 1);
});

test('Enter and Backspace in controls that hold, or stand in, other wrappers: the file that was read, byte for byte', async () => {
  const S = (tag: string, id: string, inner: string): string => `<w:sdt><w:sdtPr><w:tag w:val="${tag}"/><w:id w:val="${id}"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
  const r = (text: string): string => `<w:r><w:t>${text}</w:t></w:r>`;
  for (const [what, inner, offset] of [
    ['a control in a link', `<w:hyperlink w:anchor="z">${S('a', '5', r('ABCD'))}</w:hyperlink>`, 3],
    ['a link in a control', S('a', '5', `<w:hyperlink w:anchor="z">${r('ABCD')}</w:hyperlink>`), 3],
    ['three controls one in the other', S('a', '5', S('b', '6', S('c', '7', r('ABCD')))), 3],
    ['Enter in the outer part of two controls', S('a', '5', r('xy') + S('b', '6', r('ABCD'))), 2],
  ] as [string, string, number][]) {
    const { before, after } = await edited({ body: p(r('s') + inner + r('t')) }, '/word/document.xml', [(doc) => splitParagraph(doc, at('body', [0], offset)), (doc) => deleteBackward(doc, at('body', [1], 0))]);
    assert.equal(after, before, what);
  }
});

test('two controls of the file in two paragraphs stay two when the paragraphs are joined: by Backspace, by deleting the range between, by accepting or rejecting the mark', async () => {
  const one = '<w:sdt><w:sdtPr><w:tag w:val="f"/><w:id w:val="5"/></w:sdtPr><w:sdtContent><w:r><w:t>AAA</w:t></w:r></w:sdtContent></w:sdt>';
  const ins = '<w:pPr><w:rPr><w:ins w:id="90" w:author="Other" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr>';
  const del = '<w:pPr><w:rPr><w:del w:id="91" w:author="Other" w:date="2026-01-01T00:00:00Z"/></w:rPr></w:pPr>';
  const cases: [string, string, (doc: DocModel) => DocEdit][] = [
    ['Backspace', p(one) + p(B), (doc) => deleteBackward(doc, at('body', [1], 0))],
    ['the mark deleted', p(one) + p(B), (doc) => deleteRange(doc, { anchor: at('body', [0], 3), focus: at('body', [1], 0) })],
    ['from the middle of one to the middle of the other', p(one) + p(B), (doc) => deleteRange(doc, { anchor: at('body', [0], 2), focus: at('body', [1], 1) })],
    ['an inserted mark rejected', `<w:p>${ins}${one}</w:p>` + p(B), (doc) => rejectAll(doc)],
    ['a deleted mark accepted', `<w:p>${del}${one}</w:p>` + p(B), (doc) => acceptAll(doc)],
  ];
  for (const [what, body, edit] of cases) {
    const { after } = await edited({ body }, '/word/document.xml', [edit]);
    assert.equal(count(after, /<w:sdt>/g), 2, what);
    assert.equal(count(after, /<w:p[ >]/g), 1, `${what}: the paragraphs are one`);
  }
});
