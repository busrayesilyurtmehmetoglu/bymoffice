// Editing a document that tracks its changes: what is typed becomes a tracked
// insertion, what is deleted stays as a tracked deletion, and what cannot be
// tracked is refused instead of being changed without a trace.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import { findWrapper, paragraphText, type DocModel, type Paragraph, type Wrapper } from '../src/model/doc/types';
import {
  DocHistory,
  TRACK_UNSUPPORTED,
  deleteBackward,
  deleteColumn,
  deleteForward,
  deleteRange,
  deleteRow,
  insertBlock,
  insertColumn,
  insertHyperlink,
  insertImage,
  insertPageBreak,
  insertRow,
  insertTable,
  insertText,
  nextRevisionId,
  revisionDate,
  setTracking,
  splitParagraph,
  storyIds,
  type DocEdit,
  type DocEditOptions,
  type DocPosition,
  type DocRange,
} from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { parseXml } from '../src/xml';
import { STYLES, checkInvariants, dice, docPicture, p, paragraphAt, paragraphPaths, pos, r, randomPosition, range, run, tbl, tc, texts, tr } from './doc-ops-fixtures';
import { buildDocx as buildSpec, type DocxSpec } from './docx-fixtures';
import { schemaOrderViolations } from './docx-write-helpers.test';

const decoder = new TextDecoder();
const ME: DocEditOptions = { author: 'Me', date: '2026-10-05T10:00:00Z' };
const TRACK = '<w:trackRevisions/>';
const SECT = '<w:sectPr><w:pgSz w:w="9000" w:h="9000"/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr>';

interface Opened {
  pkg: OpcPackage;
  doc: DocModel;
  history: DocHistory<DocRange | undefined>;
  bytes: Uint8Array;
}

async function open(body: string, extra: Partial<DocxSpec> = {}): Promise<Opened> {
  const bytes = await buildSpec({ body: body + SECT, styles: STYLES, settings: TRACK, ...extra });
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
  return text.slice(text.indexOf('<w:body>') + 8, text.indexOf('<w:sectPr>'));
}

const ATTRS = ' w:author="Me" w:date="2026-10-05T10:00:00Z"';
const ins = (id: number, inner: string) => `<w:ins w:id="${id}"${ATTRS}>${inner}</w:ins>`;
const del = (id: number, inner: string) => `<w:del w:id="${id}"${ATTRS}>${inner}</w:del>`;
const t = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const dt = (text: string) => `<w:r><w:delText>${text}</w:delText></w:r>`;

const refused = (make: () => DocEdit): void => {
  assert.throws(make, (error: unknown) => error instanceof RangeError && error.message === TRACK_UNSUPPORTED);
};

/** Every w:id of a revision, bookmark or comment in the given XML texts. */
function ids(...xml: string[]): string[] {
  const out: string[] = [];
  for (const text of xml) for (const match of text.matchAll(/<w:(?:ins|del|bookmarkStart|comment|moveFrom|moveTo)\b[^>]*? w:id="(\d+)"/g)) out.push(match[1]);
  return out;
}

test('typed text is a tracked insertion, and typing on continues it', async () => {
  const o = await open(p('abc'));
  let at = run(o.history, insertText(o.doc, pos(0, 3), 'X', undefined, ME));
  assert.deepEqual(texts(o.doc), ['abcX']);
  at = run(o.history, insertText(o.doc, at, 'Y', undefined, { author: 'Me', date: '2026-10-05T11:00:00Z' }));
  assert.deepEqual(at.focus, pos(0, 5));
  // One insertion, with the date of its first character.
  assert.equal(await body(o), `<w:p>${t('abc')}${ins(1, t('XY'))}</w:p>`);
  assert.deepEqual(schemaOrderViolations(await written(o)), []);
});

test('text typed before an insertion of the same author joins it; another author gets an insertion of their own', async () => {
  const o = await open(p('abc'));
  run(o.history, insertText(o.doc, pos(0, 1), 'X', undefined, ME));
  run(o.history, insertText(o.doc, pos(0, 1), 'W', undefined, ME));
  assert.equal(await body(o), `<w:p>${t('a')}${ins(1, t('WX'))}${t('bc')}</w:p>`);
  run(o.history, insertText(o.doc, pos(0, 3), 'Q', undefined, { author: 'You', date: '2026-10-05T12:00:00Z' }));
  assert.equal(await body(o), `<w:p>${t('a')}${ins(1, t('WX'))}<w:ins w:id="2" w:author="You" w:date="2026-10-05T12:00:00Z">${t('Q')}</w:ins>${t('bc')}</w:p>`);
});

test('text typed inside the insertion of another author cuts it in two', async () => {
  const o = await open(p(`<w:ins w:id="7" w:author="You" w:date="2020-01-01T00:00:00Z">${r('abcd')}</w:ins>`));
  run(o.history, insertText(o.doc, pos(0, 2), 'X', undefined, ME));
  assert.deepEqual(texts(o.doc), ['abXcd']);
  const you = (id: number, inner: string) => `<w:ins w:id="${id}" w:author="You" w:date="2020-01-01T00:00:00Z">${inner}</w:ins>`;
  assert.equal(await body(o), `<w:p>${you(7, t('ab'))}${ins(8, t('X'))}${you(9, t('cd'))}</w:p>`);
});

test('a document that does not track is edited as before, unless the command is told to track', async () => {
  const o = await open(p('abc'), { settings: '' });
  run(o.history, insertText(o.doc, pos(0, 3), 'X'));
  run(o.history, deleteBackward(o.doc, pos(0, 1)));
  assert.equal(await body(o), `<w:p>${t('bcX')}</w:p>`);
  run(o.history, insertText(o.doc, pos(0, 0), 'Z', undefined, { ...ME, trackChanges: true }));
  assert.equal(await body(o), `<w:p>${ins(1, t('Z'))}${t('bcX')}</w:p>`);
  // False does not turn tracking off where the document asks for it.
  const tracked = await open(p('abc'));
  run(tracked.history, insertText(tracked.doc, pos(0, 0), 'Z', undefined, { ...ME, trackChanges: false }));
  assert.equal(await body(tracked), `<w:p>${ins(1, t('Z'))}${t('abc')}</w:p>`);
});

test('an enforced protection for tracked changes tracks, whatever the setting says', async () => {
  const o = await open(p('abc'), { settings: '<w:documentProtection w:edit="trackedChanges" w:enforcement="1"/>' });
  run(o.history, insertText(o.doc, pos(0, 0), 'Z', undefined, ME));
  assert.equal(await body(o), `<w:p>${ins(1, t('Z'))}${t('abc')}</w:p>`);
});

test('Backspace keeps the text as a tracked deletion; the next one joins it', async () => {
  const o = await open(p('abcd'));
  let at = run(o.history, deleteBackward(o.doc, pos(0, 4), 'char', ME));
  assert.deepEqual(at.focus, pos(0, 3));
  assert.deepEqual(texts(o.doc), ['abc']);
  at = run(o.history, deleteBackward(o.doc, at.focus, 'char', ME));
  assert.deepEqual(at.focus, pos(0, 2));
  assert.equal(await body(o), `<w:p>${t('ab')}${del(1, dt('cd'))}</w:p>`);
  assert.deepEqual(schemaOrderViolations(await written(o)), []);
});

test('Delete keeps the text as a tracked deletion and passes over what is deleted already', async () => {
  const o = await open(p('abcd'));
  let at = run(o.history, deleteForward(o.doc, pos(0, 0), 'char', ME));
  at = run(o.history, deleteForward(o.doc, at.focus, 'char', ME));
  assert.deepEqual(at.focus, pos(0, 0));
  assert.deepEqual(texts(o.doc), ['cd']);
  assert.equal(await body(o), `<w:p>${del(1, dt('ab'))}${t('cd')}</w:p>`);
  // Backspace at the end of the text reaches over the deletion to nothing: there is no text before it.
  const none = deleteBackward(o.doc, pos(0, 0), 'char', ME);
  assert.equal(none.ops.length, 0);
});

test('deleting what the same author inserted takes it out for good', async () => {
  const o = await open(p('abc'));
  const before = docPicture(o.doc);
  const at = run(o.history, insertText(o.doc, pos(0, 3), 'XY', undefined, ME));
  run(o.history, deleteBackward(o.doc, at.focus, 'char', ME));
  assert.equal(await body(o), `<w:p>${t('abc')}${ins(1, t('X'))}</w:p>`);
  run(o.history, deleteBackward(o.doc, pos(0, 4), 'char', ME));
  assert.equal(await body(o), `<w:p>${t('abc')}</w:p>`);
  assert.deepEqual(texts(o.doc), ['abc']);
  while (o.history.canUndo) o.history.undo();
  assert.equal(docPicture(o.doc), before);
});

test('deleting in the insertion of another author makes a deletion inside it', async () => {
  const o = await open(p(`<w:ins w:id="7" w:author="You" w:date="2020-01-01T00:00:00Z">${r('abcd')}</w:ins>`));
  run(o.history, deleteRange(o.doc, range(pos(0, 1), pos(0, 3)), ME));
  assert.deepEqual(texts(o.doc), ['ad']);
  assert.equal(await body(o), `<w:p><w:ins w:id="7" w:author="You" w:date="2020-01-01T00:00:00Z">${t('a')}${del(8, dt('bc'))}${t('d')}</w:ins></w:p>`);
});

test('text in a deletion is never deleted again', async () => {
  const o = await open(p(`${r('ab')}<w:del w:id="3" w:author="You" w:date="2020-01-01T00:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del>${r('cd')}`));
  assert.deepEqual(texts(o.doc), ['abcd']);
  // Backspace after "c" takes "c"; the next one passes the deletion and takes "b".
  let at = run(o.history, deleteBackward(o.doc, pos(0, 3), 'char', ME));
  at = run(o.history, deleteBackward(o.doc, at.focus, 'char', ME));
  assert.deepEqual(texts(o.doc), ['ad']);
  const xml = await body(o);
  assert.equal(xml, `<w:p>${t('a')}${del(5, dt('b'))}<w:del w:id="3" w:author="You" w:date="2020-01-01T00:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del>${del(4, dt('c'))}${t('d')}</w:p>`);
  // A range over all of it leaves the old deletion as it is.
  run(o.history, deleteRange(o.doc, range(pos(0, 0), pos(0, 2)), ME));
  assert.equal((await body(o)).split('gone').length, 2);
  assert.deepEqual(texts(o.doc), ['']);
});

test('typing over a selection deletes it and inserts after it', async () => {
  const o = await open(p('abcd'));
  const at = run(o.history, insertText(o.doc, range(pos(0, 1), pos(0, 3)), 'X', undefined, ME));
  assert.deepEqual(at.focus, pos(0, 2));
  assert.deepEqual(texts(o.doc), ['aXd']);
  assert.equal(await body(o), `<w:p>${t('a')}${del(1, dt('bc'))}${ins(2, t('X'))}${t('d')}</w:p>`);
});

test('formatted text keeps its formatting in the deletion', async () => {
  const o = await open(p(r('ab', '<w:b/>') + r('cd')));
  run(o.history, deleteRange(o.doc, range(pos(0, 1), pos(0, 3)), ME));
  assert.equal(
    await body(o),
    `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>a</w:t></w:r>${del(1, '<w:r><w:rPr><w:b/></w:rPr><w:delText>b</w:delText></w:r>' + dt('c'))}${t('d')}</w:p>`,
  );
});

test('Enter marks the paragraph mark it makes as inserted; Backspace by the same author takes it back', async () => {
  const o = await open(p('abcd', '<w:jc w:val="center"/>'));
  const before = docPicture(o.doc);
  const at = run(o.history, splitParagraph(o.doc, pos(0, 2), ME));
  assert.deepEqual(at.focus, pos(1, 0));
  assert.deepEqual(texts(o.doc), ['ab', 'cd']);
  const mark = paragraphAt(o.doc, 0).props.markProps;
  assert.deepEqual(mark && mark.inserted, { id: 1, author: 'Me', date: '2026-10-05T10:00:00Z' });
  assert.equal(paragraphAt(o.doc, 1).props.markProps, undefined);
  assert.equal(await body(o), `<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:ins w:id="1"${ATTRS}/></w:rPr></w:pPr>${t('ab')}</w:p><w:p><w:pPr><w:jc w:val="center"/></w:pPr>${t('cd')}</w:p>`);
  assert.deepEqual(schemaOrderViolations(await written(o)), []);
  const joined = run(o.history, deleteBackward(o.doc, at.focus, 'char', ME));
  assert.deepEqual(joined.focus, pos(0, 2));
  assert.deepEqual(texts(o.doc), ['abcd']);
  assert.equal(await body(o), `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${t('abcd')}</w:p>`);
  while (o.history.canUndo) o.history.undo();
  assert.equal(docPicture(o.doc), before);
});

test('Enter at the end of a heading does not restyle a paragraph without a trace', async () => {
  const o = await open(p('Title', '<w:pStyle w:val="Heading1"/>'));
  run(o.history, splitParagraph(o.doc, pos(0, 5), ME));
  assert.equal(paragraphAt(o.doc, 0).props.style, 'Heading1');
  assert.equal(paragraphAt(o.doc, 1).props.style, 'Heading1');
  assert.ok(paragraphAt(o.doc, 0).props.markProps?.inserted);
});

test('joining two paragraphs marks the paragraph mark as deleted and keeps both paragraphs', async () => {
  const o = await open(p('ab') + p('cd'));
  let at = run(o.history, deleteBackward(o.doc, pos(1, 0), 'char', ME));
  assert.deepEqual(texts(o.doc), ['ab', 'cd']);
  assert.deepEqual(at.focus, pos(0, 2));
  assert.equal(await body(o), `<w:p><w:pPr><w:rPr><w:del w:id="1"${ATTRS}/></w:rPr></w:pPr>${t('ab')}</w:p>${p('cd')}`);
  // Again: nothing to change, the caret still moves on.
  const again = deleteBackward(o.doc, pos(1, 0), 'char', ME);
  assert.equal(again.ops.length, 0);
  assert.deepEqual(again.selection.focus, pos(0, 2));
  // Delete at the end of the first paragraph: the same mark, the caret goes to the next paragraph.
  const forward = deleteForward(o.doc, pos(0, 2), 'char', ME);
  assert.equal(forward.ops.length, 0);
  assert.deepEqual(forward.selection.focus, pos(1, 0));
  at = run(o.history, deleteBackward(o.doc, at.focus, 'char', ME));
  assert.deepEqual(texts(o.doc), ['a', 'cd']);
});

test('Delete at the end of a paragraph marks its mark deleted and moves to the next paragraph', async () => {
  const o = await open(p('ab') + p('cd'));
  const at = run(o.history, deleteForward(o.doc, pos(0, 2), 'char', ME));
  assert.deepEqual(at.focus, pos(1, 0));
  assert.ok(paragraphAt(o.doc, 0).props.markProps?.deleted);
  // Before a table nothing happens, as without tracking.
  const table = await open(p('ab') + tbl(tr(tc('x')), [2000]) + p('z'));
  assert.equal(deleteForward(table.doc, pos(0, 2), 'char', ME).ops.length, 0);
});

test('Backspace at the start of a list item is a tracked join, not an untracked change of the list', async () => {
  const o = await open(p('one') + p('two', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'), {
    numbering: '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>',
  });
  run(o.history, deleteBackward(o.doc, pos(1, 0), 'char', ME));
  assert.deepEqual(paragraphAt(o.doc, 1).props.numbering, { numId: 1, level: 0 });
  assert.ok(paragraphAt(o.doc, 0).props.markProps?.deleted);
});

test('a range over several paragraphs: the text and the marks between are deleted, the paragraphs stay', async () => {
  const o = await open(p('abcd') + p('efgh') + p('ijkl'));
  const at = run(o.history, deleteRange(o.doc, range(pos(2, 2), pos(0, 2)), ME));
  assert.deepEqual(at.focus, pos(0, 2));
  assert.deepEqual(texts(o.doc), ['ab', '', 'kl']);
  const mark = (id: number) => `<w:pPr><w:rPr><w:del w:id="${id}"${ATTRS}/></w:rPr></w:pPr>`;
  assert.equal(await body(o), `<w:p>${mark(2)}${t('ab')}${del(1, dt('cd'))}</w:p><w:p>${mark(4)}${del(3, dt('efgh'))}</w:p><w:p>${del(5, dt('ij'))}${t('kl')}</w:p>`);
  assert.deepEqual(schemaOrderViolations(await written(o)), []);
});

test('a range over paragraphs the same author made takes their marks and text out for good', async () => {
  const o = await open(p('abcd'));
  run(o.history, insertText(o.doc, pos(0, 2), 'X\nY\nZ', undefined, ME));
  assert.deepEqual(texts(o.doc), ['abX', 'Y', 'Zcd']);
  for (const index of [0, 1]) assert.equal(paragraphAt(o.doc, index).props.markProps?.inserted?.author, 'Me');
  assert.equal(paragraphAt(o.doc, 2).props.markProps, undefined);
  const xml = await body(o);
  assert.equal(new Set(ids(xml)).size, ids(xml).length, xml);
  run(o.history, deleteRange(o.doc, range(pos(0, 2), pos(2, 1)), ME));
  assert.deepEqual(texts(o.doc), ['abcd']);
  assert.equal(await body(o), `<w:p>${t('abcd')}</w:p>`);
});

test('pasted lines are inserted paragraphs', async () => {
  const o = await open(p('ab'));
  const at = run(o.history, insertText(o.doc, pos(0, 1), 'x\ny', undefined, ME));
  assert.deepEqual(at.focus, pos(1, 1));
  assert.equal(await body(o), `<w:p><w:pPr><w:rPr><w:ins w:id="2"${ATTRS}/></w:rPr></w:pPr>${t('a')}${ins(1, t('x'))}</w:p><w:p>${ins(3, t('y'))}${t('b')}</w:p>`);
});

test('a page break is a tracked insertion', async () => {
  const o = await open(p('ab'));
  run(o.history, insertPageBreak(o.doc, pos(0, 1), ME));
  const xml = await body(o);
  assert.equal(xml, `<w:p><w:pPr><w:rPr><w:ins w:id="2"${ATTRS}/></w:rPr></w:pPr>${t('a')}${ins(1, '<w:r><w:br w:type="page"/></w:r>')}</w:p><w:p>${t('b')}</w:p>`);
});

test('a field goes into the deletion whole, its code as deleted code', async () => {
  const field = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>today</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
  const o = await open(p(r('a') + field + r('b')));
  assert.deepEqual(texts(o.doc), ['atodayb']);
  run(o.history, deleteRange(o.doc, range(pos(0, 1), pos(0, 6)), ME));
  assert.deepEqual(texts(o.doc), ['ab']);
  const xml = await body(o);
  assert.ok(/^<w:p><w:r><w:t[^>]*>a<\/w:t><\/w:r><w:del w:id="1"[^>]*>.*<\/w:del><w:r><w:t[^>]*>b<\/w:t><\/w:r><\/w:p>$/.test(xml), xml);
  assert.ok(xml.indexOf('<w:delInstrText xml:space="preserve"> DATE </w:delInstrText>') > 0, xml);
  assert.ok(xml.indexOf('<w:delText>today</w:delText>') > 0, xml);
  assert.equal(xml.indexOf('<w:instrText'), -1);
  assert.equal(xml.split('<w:fldChar').length, 4);
});

test('a note reference stays, deleted, and its note stays with it', async () => {
  const o = await open(p(r('a') + '<w:r><w:footnoteReference w:id="2"/></w:r>' + r('b')), {
    footnotes: '<w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>note</w:t></w:r></w:p></w:footnote>',
  });
  run(o.history, deleteRange(o.doc, range(pos(0, 0), pos(0, 2)), ME));
  assert.ok(o.doc.footnotes && o.doc.footnotes.items.has(2));
  const xml = await body(o);
  assert.ok(xml.indexOf('<w:footnoteReference w:id="2"/>') > xml.indexOf('<w:del '), xml);
});

test('a picture, a tab and a line break go into the deletion with their run; Backspace takes a picture as one unit', async () => {
  const drawing = '<w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="1" name="p"/></wp:inline></w:drawing></w:r>';
  const o = await open(p(r('a') + drawing + '<w:r><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t></w:r>'));
  assert.deepEqual(texts(o.doc), ['a\tb\nc']);
  // After the picture: Backspace takes the picture alone.
  run(o.history, deleteBackward(o.doc, pos(0, 1), 'char', ME));
  assert.deepEqual(texts(o.doc), ['a\tb\nc']);
  let xml = await body(o);
  assert.ok(/<w:del w:id="1"[^>]*><w:r><w:drawing>/.test(xml), xml);
  assert.equal(xml.split('<w:drawing>').length, 2);
  // It is not a unit any more: the next Backspace reaches the text before it.
  run(o.history, deleteBackward(o.doc, pos(0, 1), 'char', ME));
  assert.deepEqual(texts(o.doc), ['\tb\nc']);
  run(o.history, deleteRange(o.doc, range(pos(0, 0), pos(0, 3)), ME));
  assert.deepEqual(texts(o.doc), ['c']);
  xml = await body(o);
  assert.equal(xml.split('<w:drawing>').length, 2);
  assert.ok(xml.indexOf('<w:tab/><w:delText>b</w:delText><w:br/>') > 0, xml);
  assert.equal(xml.indexOf('<w:t>b</w:t>'), -1);
  assert.deepEqual(schemaOrderViolations(await written(o)), []);
  // Every deletion of the author that touches another became one with it.
  assert.equal(xml.split('<w:del ').length, 2, xml);
});

test('a hyperlink keeps its deleted text, and text typed in a link is an insertion inside it', async () => {
  const o = await open(p(r('see ') + `<w:hyperlink r:id="rIdL">${r('the site')}</w:hyperlink>` + r(' now')), { rels: [['rIdL', 'hyperlink', 'https://example.org/', true]] });
  run(o.history, deleteRange(o.doc, range(pos(0, 4), pos(0, 8)), ME));
  run(o.history, insertText(o.doc, pos(0, 6), 'X', undefined, ME));
  assert.deepEqual(texts(o.doc), ['see siXte now']);
  const xml = await body(o);
  assert.ok(xml.indexOf(`<w:hyperlink r:id="rIdL">${del(1, '<w:r><w:delText xml:space="preserve">the </w:delText></w:r>')}<w:r><w:t>si</w:t></w:r>${ins(2, t('X'))}<w:r><w:t>te</w:t></w:r></w:hyperlink>`) > 0, xml);
});

test('the header, a footnote and a table cell are tracked like the body', async () => {
  const o = await open(tbl(tr(tc('cell')), [2000]) + p(r('a') + '<w:r><w:footnoteReference w:id="2"/></w:r>') + '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>', {
    headers: { rIdH: p('head') },
    footnotes: '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>note</w:t></w:r></w:p></w:footnote>',
  });
  run(o.history, insertText(o.doc, { story: 'body', path: [0, 0, 0, 0], offset: 4 }, '!', undefined, ME));
  run(o.history, deleteBackward(o.doc, { story: 'body', path: [0, 0, 0, 0], offset: 1 }, 'char', ME));
  run(o.history, insertText(o.doc, pos(0, 4, 'header:rIdH'), '!', undefined, ME));
  run(o.history, deleteBackward(o.doc, pos(0, 1, 'header:rIdH'), 'char', ME));
  run(o.history, insertText(o.doc, pos(0, 4, 'footnote:2'), '!', undefined, ME));
  run(o.history, deleteBackward(o.doc, pos(0, 1, 'footnote:2'), 'char', ME));
  const main = await written(o);
  const header = await written(o, '/word/header1.xml');
  const notes = await written(o, '/word/footnotes.xml');
  for (const xml of [main, header, notes]) {
    assert.equal(xml.split('<w:ins ').length, 2, xml);
    assert.equal(xml.split('<w:del ').length, 2, xml);
    assert.equal(xml.split('<w:delText').length, 2, xml);
    assert.deepEqual(schemaOrderViolations(xml), []);
    parseXml(xml);
  }
  // No id is used twice across the parts.
  const all = ids(main, header, notes);
  assert.equal(all.length, 6);
  assert.equal(new Set(all).size, 6, all.join(','));
});

test('new ids are above every id of a revision, bookmark or comment in any story', async () => {
  const o = await open(p('<w:bookmarkStart w:id="40" w:name="b"/>' + r('a') + '<w:bookmarkEnd w:id="40"/>') + p(`<w:ins w:id="12" w:author="You">${r('x')}</w:ins>`), {
    headers: { rIdH: p(`<w:del w:id="77" w:author="You"><w:r><w:delText>d</w:delText></w:r></w:del>`) },
    comments: '<w:comment w:id="90" w:author="You"><w:p/></w:comment>',
    footnotes: '<w:footnote w:id="2"><w:p><w:pPr><w:rPr><w:ins w:id="120" w:author="You"/></w:rPr></w:pPr></w:p></w:footnote>',
  });
  run(o.history, insertText(o.doc, pos(0, 1), 'Z', undefined, ME));
  assert.equal(findWrapper(paragraphAt(o.doc, 0).children.filter((inline) => inline.wrap)[0], 'ins')?.id, 121);
  assert.equal(nextRevisionId(o.doc), 122);
});

test('an undone id is not given again', async () => {
  const o = await open(p('abc'));
  run(o.history, insertText(o.doc, pos(0, 0), 'X', undefined, ME));
  o.history.undo();
  run(o.history, insertText(o.doc, pos(0, 0), 'Y', undefined, ME));
  assert.equal(await body(o), `<w:p>${ins(2, t('Y'))}${t('abc')}</w:p>`);
});

test('undo brings back the bytes that were read', async () => {
  const o = await open(p('abcd') + p('efgh') + tbl(tr(tc('x') + tc('y')), [1000, 1000]) + p('end'));
  const original = await written(o);
  const picture = docPicture(o.doc);
  run(o.history, insertText(o.doc, pos(0, 2), 'new', undefined, ME));
  run(o.history, deleteRange(o.doc, range(pos(0, 1), pos(1, 2)), ME));
  run(o.history, splitParagraph(o.doc, pos(1, 1), ME));
  run(o.history, deleteBackward(o.doc, { story: 'body', path: [3, 0, 1, 0], offset: 1 }, 'char', ME));
  assert.notEqual(await written(o), original);
  // Redo gives the same file again.
  const edited = await written(o);
  let steps = 0;
  while (o.history.canUndo) {
    o.history.undo();
    steps++;
  }
  assert.equal(steps, 4);
  assert.equal(docPicture(o.doc), picture);
  assert.equal(await written(o), original);
  while (o.history.canRedo) o.history.redo();
  assert.equal(await written(o), edited);
});

/** What a model holds of revisions, in a form that survives writing and reading. */
function revisions(doc: DocModel): string[] {
  const out: string[] = [];
  const chain = (wrap: Wrapper | undefined): string => {
    const parts: string[] = [];
    for (let w = wrap; w; w = w.parent) parts.push(w.kind === 'ins' || w.kind === 'del' ? `${w.kind}#${w.id}/${w.author}/${w.date}` : w.kind);
    return parts.join('>');
  };
  for (const story of storyIds(doc)) {
    for (const path of paragraphPaths(doc, story)) {
      const paragraph: Paragraph = paragraphAt(doc, path, story);
      const mark = paragraph.props.markProps;
      out.push(`${story}:${path.join('.')}|${paragraphText(paragraph)}|${mark && mark.inserted ? 'I' + mark.inserted.id : ''}${mark && mark.deleted ? 'D' + mark.deleted.id : ''}`);
      for (const inline of paragraph.children) {
        if (inline.type !== 'run') continue;
        out.push('  ' + inline.content.map((item) => (item.type === 'text' || item.type === 'deletedText' ? `${item.type}:${item.text}` : item.type)).join(',') + ' @' + chain(inline.wrap));
      }
    }
  }
  return out;
}

test('save, open again: the model holds the same revisions', async () => {
  const o = await open(p('abcd') + p('efgh') + p('ijkl'));
  run(o.history, insertText(o.doc, pos(0, 2), 'new\nline', undefined, ME));
  run(o.history, deleteRange(o.doc, range(pos(1, 6), pos(3, 2)), ME));
  run(o.history, insertText(o.doc, range(pos(3, 0), pos(3, 1)), 'K', undefined, ME));
  await writeDocx(o.pkg, o.doc);
  const again = await readDocx(await OpcPackage.open(await o.pkg.save()));
  assert.deepEqual(revisions(again), revisions(o.doc));
  assert.equal(again.settings.trackRevisions, true);
  checkInvariants(o.doc, false);
});

test('what cannot be tracked is refused and changes nothing', async () => {
  const o = await open(p('abcd') + tbl(tr(tc('x') + tc('y')) + tr(tc('z') + tc('w')), [1000, 1000]) + p('end'), {
    numbering: '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="-"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>',
  });
  const before = docPicture(o.doc);
  const span = range(pos(0, 1), pos(0, 3));
  const cell: DocPosition = { story: 'body', path: [1, 0, 0, 0], offset: 0 };
  // (Formatting of text, of paragraphs and lists is tracked: test/doc-track-format.test.ts.)
  refused(() => insertTable(o.doc, pos(0, 1), 2, 2));
  refused(() => insertRow(o.doc, cell, 'below'));
  refused(() => deleteRow(o.doc, cell));
  refused(() => insertColumn(o.doc, cell, 'right'));
  refused(() => deleteColumn(o.doc, cell));
  refused(() => insertHyperlink(o.doc, span, 'https://example.org/'));
  refused(() => insertImage(o.doc, pos(0, 1), { bytes: new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4]), type: 'image/png', width: 9525, height: 9525 }));
  refused(() => insertBlock(o.doc, pos(0, 1), { type: 'paragraph', props: {}, children: [] }));
  // A range that takes cells of a table, or a table whole.
  refused(() => deleteRange(o.doc, range(pos(0, 1), pos(2, 1)), ME));
  refused(() => deleteRange(o.doc, range(pos(0, 1), { story: 'body', path: [1, 0, 0, 0], offset: 1 }), ME));
  refused(() => deleteRange(o.doc, range({ story: 'body', path: [1, 0, 0, 0], offset: 0 }, { story: 'body', path: [1, 1, 1, 0], offset: 1 }), ME));
  refused(() => insertText(o.doc, range(pos(0, 1), pos(2, 1)), 'x', undefined, ME));
  assert.equal(docPicture(o.doc), before);
  assert.equal(o.doc.body.modified, undefined);
});

test('the setup of a document names the author and the clock', async () => {
  const o = await open(p('abc'), { settings: '' });
  setTracking(o.doc, { force: true, author: 'Ada', now: () => new Date(Date.UTC(2026, 0, 2, 3, 4, 5, 678)) });
  run(o.history, insertText(o.doc, pos(0, 0), 'Z'));
  assert.equal(await body(o), `<w:p><w:ins w:id="1" w:author="Ada" w:date="2026-01-02T03:04:05Z">${t('Z')}</w:ins>${t('abc')}</w:p>`);
  // Without a name: the stand-in, never an empty author.
  const plain = await open(p('abc'));
  setTracking(plain.doc, { now: () => new Date(0) });
  run(plain.history, insertText(plain.doc, pos(0, 0), 'Z'));
  assert.equal(await body(plain), `<w:p><w:ins w:id="1" w:author="Author" w:date="1970-01-01T00:00:00Z">${t('Z')}</w:ins>${t('abc')}</w:p>`);
  assert.equal(revisionDate(new Date(NaN)), '1970-01-01T00:00:00Z');
});

test('the name of an author is written as text, never as markup', async () => {
  const o = await open(p('abc'));
  run(o.history, insertText(o.doc, pos(0, 0), 'Z', undefined, { author: '"><w:evil/>&', date: ME.date }));
  const xml = await written(o);
  assert.ok(xml.indexOf('w:author="&quot;&gt;&lt;w:evil/&gt;&amp;"') > 0 || xml.indexOf('w:author="&quot;>&lt;w:evil/>&amp;"') > 0, xml);
  const again = await readDocx(await OpcPackage.open(await o.pkg.save()));
  assert.equal(findWrapper(paragraphAt(again, 0).children[0], 'ins')?.author, '"><w:evil/>&');
});

test('random tracked edits: the file stays well formed and in schema order, and undoing all of them restores the document', async () => {
  for (let seed = 1; seed <= 12; seed++) {
    const d = dice(seed * 7919);
    const o = await open(
      p('alpha beta gamma') + p(r('bold', '<w:b/>') + r(' and plain') + `<w:ins w:id="5" w:author="You">${r(' theirs')}</w:ins>`) + p('') + tbl(tr(tc('one') + tc(p('two') + p('three'))), [1500, 1500]) + p('omega end'),
      { headers: { rIdH: p('header text') } },
    );
    const picture = docPicture(o.doc);
    const original = await written(o);
    let made = 0;
    for (let step = 0; step < 40; step++) {
      const stories = storyIds(o.doc);
      const story = d.pick(stories);
      const paths = paragraphPaths(o.doc, story);
      const at = randomPosition(o.doc, d, story, paths);
      const other = randomPosition(o.doc, d, story, paths);
      const who: DocEditOptions = { author: d.chance(0.7) ? 'Me' : 'Other', date: ME.date };
      let edit: DocEdit | undefined;
      try {
        switch (d.int(6)) {
          case 0:
            edit = insertText(o.doc, at, d.pick(['x', 'two words', 'a\nb', '\t', ' ']), undefined, who);
            break;
          case 1:
            edit = deleteBackward(o.doc, at, d.chance(0.3) ? 'word' : 'char', who);
            break;
          case 2:
            edit = deleteForward(o.doc, at, d.chance(0.3) ? 'word' : 'char', who);
            break;
          case 3:
            edit = deleteRange(o.doc, range(at, other), who);
            break;
          case 4:
            edit = splitParagraph(o.doc, at, who);
            break;
          default:
            edit = insertText(o.doc, range(at, other), 'Z', undefined, who);
        }
      } catch (error) {
        // A range over a table is refused; nothing else may be.
        assert.ok(error instanceof RangeError && error.message === TRACK_UNSUPPORTED, String(error));
        continue;
      }
      if (edit.ops.length === 0) continue;
      run(o.history, edit);
      made++;
      checkInvariants(o.doc, false);
    }
    assert.ok(made > 10, `seed ${seed}: ${made} edits`);
    const xml = await written(o);
    parseXml(xml);
    assert.deepEqual(schemaOrderViolations(xml), [], `seed ${seed}`);
    // No insertion inside an insertion, no deletion inside a deletion, and no text that is not deleted text in a deletion.
    const root = parseXml(xml).root;
    const walk = (el: typeof root, inIns: boolean, inDel: boolean): void => {
      const name = el.name;
      assert.ok(!(name === 'w:ins' && inIns && el.children.length > 0), `seed ${seed}: nested insertion`);
      assert.ok(!(name === 'w:del' && inDel && el.children.length > 0), `seed ${seed}: nested deletion`);
      assert.ok(!(name === 'w:t' && inDel), `seed ${seed}: w:t in a deletion`);
      assert.ok(!(name === 'w:delText' && !inDel), `seed ${seed}: w:delText outside a deletion`);
      for (const child of el.children) if (child.type === 'element') walk(child, inIns || (name === 'w:ins' && el.children.length > 0), inDel || (name === 'w:del' && el.children.length > 0));
    };
    walk(root, false, false);
    const all = ids(xml);
    assert.equal(new Set(all).size, all.length, `seed ${seed}: ${all.join(',')}`);
    // The text that is there or deleted still holds every character the document had.
    while (o.history.canUndo) o.history.undo();
    assert.equal(docPicture(o.doc), picture, `seed ${seed}`);
    assert.equal(await written(o), original, `seed ${seed}`);
  }
});

// ------------------------------------------------------- findings of the review

const FIELD =
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>today</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
const NOTES = '<w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>note</w:t></w:r></w:p></w:footnote>';
const DRAWING = '<w:r><w:drawing><wp:inline><wp:extent cx="9525" cy="9525"/><wp:docPr id="1" name="p"/></wp:inline></w:drawing></w:r>';
const MATH = '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x+1</m:t></m:r></m:oMath>';
const OWN = (author: string, inner: string) => `<w:ins w:id="7" w:author="${author}" w:date="2020-01-01T00:00:00Z">${inner}</w:ins>`;

test('of what the same author inserted only text goes for good: a note reference, a field, a picture and an equation stay, deleted', async () => {
  const o = await open(p(r('a') + OWN('Me', r('x') + '<w:r><w:footnoteReference w:id="2"/></w:r>' + FIELD + DRAWING + MATH + '<w:r><w:tab/><w:br/><w:br w:type="page"/><w:t>y</w:t></w:r>') + r('b')), { footnotes: NOTES });
  const original = await written(o);
  const picture = docPicture(o.doc);
  const length = texts(o.doc)[0].length;
  run(o.history, deleteRange(o.doc, range(pos(0, 0), pos(0, length)), ME));
  assert.deepEqual(texts(o.doc), ['']);
  const xml = await body(o);
  // The note is still there, and so is everything that is more than text.
  assert.ok(o.doc.footnotes && o.doc.footnotes.items.has(2));
  for (const kept of ['<w:footnoteReference w:id="2"/>', '<w:delInstrText xml:space="preserve"> DATE </w:delInstrText>', '<w:delText>today</w:delText>', '<w:drawing>', '<m:oMath', '<w:br w:type="page"/>']) {
    assert.equal(xml.split(kept).length, 2, `${kept}\n${xml}`);
  }
  assert.equal(xml.split('<w:fldChar').length, 4, xml);
  // The text the author typed is gone: no "x", no "y", no tab.
  assert.equal(xml.indexOf('>x<'), -1, xml);
  assert.equal(xml.indexOf('>y<'), -1, xml);
  assert.equal(xml.indexOf('<w:tab/>'), -1, xml);
  // What stays of the insertion is deleted inside it.
  assert.ok(/<w:ins w:id="7"[^>]*>(<w:del [^>]*>.*<\/w:del>)+<\/w:ins>/.test(xml), xml);
  assert.deepEqual(schemaOrderViolations(await written(o)), []);
  o.history.undo();
  assert.equal(docPicture(o.doc), picture);
  assert.equal(await written(o), original);
});

test('Backspace over a note reference of the same author keeps the note', async () => {
  const o = await open(p(r('a') + OWN('Me', r('x') + '<w:r><w:footnoteReference w:id="2"/></w:r>')), { footnotes: NOTES });
  run(o.history, deleteBackward(o.doc, pos(0, 2), 'char', ME));
  assert.ok(o.doc.footnotes && o.doc.footnotes.items.has(2));
  const xml = await body(o);
  assert.ok(/<w:del [^>]*><w:r><w:footnoteReference w:id="2"\/><\/w:r><\/w:del>/.test(xml), xml);
});

test('an editor without a name owns no insertion: what "Author" inserted is deleted as a tracked change', async () => {
  const o = await open(p(r('a') + OWN('Author', r('xy')) + r('b')) + p(''));
  setTracking(o.doc, { now: () => new Date(0) });
  run(o.history, insertText(o.doc, pos(0, 3), 'Z'));
  run(o.history, deleteRange(o.doc, range(pos(0, 0), pos(0, 5))));
  assert.deepEqual(texts(o.doc), ['', '']);
  const xml = await body(o);
  // "Z" continued the insertion (the same name is written), and all of it is deleted, not taken out.
  assert.equal(xml.split('<w:delText>xyZ</w:delText>').length, 2, xml);
  // Nor a paragraph mark: Enter, then Backspace, leaves both marks of the change.
  const at = run(o.history, splitParagraph(o.doc, pos(1, 0)));
  run(o.history, deleteBackward(o.doc, at.focus));
  assert.equal(texts(o.doc).length, 3);
  assert.ok(paragraphAt(o.doc, 1).props.markProps?.inserted && paragraphAt(o.doc, 1).props.markProps?.deleted);
  // The name that is written is still the stand-in.
  assert.ok(xml.indexOf('w:author="Author" w:date="1970-01-01T00:00:00Z"') > 0, xml);
});

test('text typed into a content control inside the insertion of somebody else is refused: the control would be written three times', async () => {
  const sdt = (inner: string) => `<w:sdt><w:sdtPr><w:id w:val="99"/><w:dataBinding w:xpath="/a"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
  const o = await open(p(OWN('You', sdt(r('abcd')))) + p(OWN('You', `<w:hyperlink w:anchor="here">${r('link')}</w:hyperlink>`)) + p(OWN('Me', sdt(r('mine')))) + p(sdt(OWN('You', r('inside')))));
  const before = docPicture(o.doc);
  refused(() => insertText(o.doc, pos(0, 2), 'X', undefined, ME));
  refused(() => insertText(o.doc, pos(0, 2), 'X\nY', undefined, ME));
  refused(() => insertPageBreak(o.doc, pos(0, 2), ME));
  // Several paragraphs into a control inside the author's own insertion: each would get a copy of the control.
  refused(() => insertText(o.doc, pos(2, 2), 'X\nY', undefined, ME));
  assert.equal(docPicture(o.doc), before);
  // One line there continues the insertion; a link is cut as before; a control around the insertion is not cut at all.
  run(o.history, insertText(o.doc, pos(2, 2), 'X', undefined, ME));
  run(o.history, insertText(o.doc, pos(1, 2), 'X', undefined, ME));
  run(o.history, insertText(o.doc, pos(3, 2), 'X', undefined, ME));
  const xml = await body(o);
  assert.equal(xml.split('<w:sdt>').length, 4, xml);
  assert.equal(xml.split('<w:hyperlink ').length, 4, xml);
  assert.deepEqual(texts(o.doc), ['abcd', 'liXnk', 'miXne', 'inXside']);
});

test('no id is written twice: a deletion text is typed into, and the kept change of a paragraph that is split', async () => {
  const gone = '<w:del w:id="3" w:author="You"><w:r><w:delText>g1</w:delText></w:r><w:bookmarkStart w:id="30" w:name="b"/><w:r><w:delText>g2</w:delText></w:r><w:bookmarkEnd w:id="30"/></w:del>';
  const change = '<w:jc w:val="center"/><w:rPr><w:b/><w:rPrChange w:id="41" w:author="You"><w:rPr/></w:rPrChange></w:rPr><w:pPrChange w:id="40" w:author="You"><w:pPr/></w:pPrChange>';
  const o = await open(p(r('ab') + gone + r('cd')) + p('efgh', change));
  run(o.history, insertText(o.doc, pos(0, 2), 'X', undefined, ME));
  run(o.history, splitParagraph(o.doc, pos(1, 2), ME));
  run(o.history, insertText(o.doc, pos(2, 1), 'p\nq', undefined, ME));
  run(o.history, insertPageBreak(o.doc, pos(3, 1), ME));
  const xml = await body(o);
  const all = Array.from(xml.matchAll(/<w:(?:ins|del|pPrChange|rPrChange)\b[^>]*? w:id="(\d+)"/g)).map((match) => match[1]);
  assert.equal(new Set(all).size, all.length, all.join(',') + '\n' + xml);
  assert.equal(xml.split('<w:pPrChange ').length, 5, xml);
  assert.equal(xml.split('g1').length, 2);
  assert.equal(xml.split('g2').length, 2);
  assert.deepEqual(schemaOrderViolations(await written(o)), []);
  while (o.history.canUndo) o.history.undo();
});

test('a deleted equation goes into the deletion', async () => {
  const o = await open(p(r('a') + MATH + r('b')));
  assert.deepEqual(texts(o.doc), ['ax+1b']);
  run(o.history, deleteRange(o.doc, range(pos(0, 1), pos(0, 4)), ME));
  assert.deepEqual(texts(o.doc), ['ab']);
  const xml = await body(o);
  assert.ok(/<w:del w:id="1"[^>]*><m:oMath[^>]*>.*<\/m:oMath><\/w:del>/.test(xml), xml);
  parseXml(await written(o));
});

test('the ids of a header, of a footer and of the comments count each by themselves', async () => {
  const base = { headers: undefined, footers: undefined, comments: undefined } as Partial<DocxSpec>;
  const cases: [Partial<DocxSpec>, number][] = [
    [{ ...base, headers: { rIdH: p('<w:bookmarkStart w:id="300" w:name="h"/><w:bookmarkEnd w:id="300"/>') } }, 301],
    [{ ...base, footers: { rIdF: p(`<w:ins w:id="400" w:author="You">${r('f')}</w:ins>`) } }, 401],
    [{ ...base, comments: '<w:comment w:id="500" w:author="You"><w:p/></w:comment>' }, 501],
    [{ ...base, endnotes: '<w:endnote w:id="1"><w:p><w:pPr><w:rPr><w:del w:id="600" w:author="You"/></w:rPr></w:pPr></w:p></w:endnote>' }, 601],
  ];
  for (const [extra, expected] of cases) {
    const o = await open(p('abc'), extra);
    assert.equal(nextRevisionId(o.doc), expected);
  }
});

test('a date that is not ISO 8601 is refused', async () => {
  const o = await open(p('abc'));
  for (const date of ['yesterday', '2026-10-05', '2026-10-05T10:00:00Z" w:x="1', '05.10.2026 10:00', '2026-13-45T99:00:00Z']) {
    assert.throws(() => insertText(o.doc, pos(0, 0), 'X', undefined, { author: 'Me', date }), RangeError, date);
  }
  for (const date of ['2026-10-05T10:00:00Z', '2026-10-05T10:00:00', '2026-10-05T10:00:00.123+03:00']) insertText(o.doc, pos(0, 0), 'X', undefined, { author: 'Me', date });
});

test('random tracked edits in a document with a note, a field, a picture, a control and an equation: nothing but typed text ever leaves', async () => {
  const sdt = `<w:sdt><w:sdtPr><w:id w:val="99"/></w:sdtPr><w:sdtContent>${r('control')}</w:sdtContent></w:sdt>`;
  /** What must never go: counted in every story, in the body of the file. */
  const kept = ['<w:footnoteReference ', '<w:drawing>', '<m:oMath', '<w:sdt>', 'DATE'];
  for (let seed = 1; seed <= 10; seed++) {
    const d = dice(seed * 104729);
    const o = await open(
      p(r('alpha ') + '<w:r><w:footnoteReference w:id="2"/></w:r>' + r(' beta')) + p(r('one ') + FIELD + r(' two ') + DRAWING + r(' three')) + p(r('x ') + MATH + r(' y ') + sdt + r(' z')) + p(OWN('Me', r('mine ') + DRAWING + r(' too'))) + p('omega end'),
      { footnotes: NOTES },
    );
    const picture = docPicture(o.doc);
    const original = await written(o);
    const count = (xml: string) => kept.map((mark) => xml.split(mark).length - 1);
    const had = count(original);
    let made = 0;
    for (let step = 0; step < 50; step++) {
      const story = d.pick(storyIds(o.doc));
      const paths = paragraphPaths(o.doc, story);
      const at = randomPosition(o.doc, d, story, paths);
      const other = randomPosition(o.doc, d, story, paths);
      const who: DocEditOptions = { author: d.chance(0.8) ? 'Me' : 'Other', date: ME.date };
      let edit: DocEdit | undefined;
      try {
        switch (d.int(7)) {
          case 0:
            edit = insertText(o.doc, at, d.pick(['x', 'two words', 'a\nb', '\t', '\v']), undefined, who);
            break;
          case 1:
            edit = deleteBackward(o.doc, at, d.chance(0.3) ? 'word' : 'char', who);
            break;
          case 2:
            edit = deleteForward(o.doc, at, d.chance(0.3) ? 'word' : 'char', who);
            break;
          case 3:
            edit = deleteRange(o.doc, range(at, other), who);
            break;
          case 4:
            edit = splitParagraph(o.doc, at, who);
            break;
          case 5:
            edit = insertPageBreak(o.doc, at, who);
            break;
          default:
            edit = insertText(o.doc, range(at, other), 'Z', undefined, who);
        }
      } catch (error) {
        // Refused: a control that would be copied; a page break in a note is refused as it always is.
        assert.ok(error instanceof RangeError, String(error));
        continue;
      }
      if (edit.ops.length === 0) continue;
      run(o.history, edit);
      made++;
      checkInvariants(o.doc, false);
    }
    assert.ok(made > 10, `seed ${seed}: ${made} edits`);
    const xml = await written(o);
    parseXml(xml);
    assert.deepEqual(schemaOrderViolations(xml), [], `seed ${seed}`);
    // Enter inside the control makes it two (as without tracking); nothing else may change in number.
    const now = count(xml);
    assert.ok(now[3] >= had[3], `seed ${seed}`);
    now[3] = had[3];
    assert.deepEqual(now, had, `seed ${seed}\n${xml}`);
    assert.ok(o.doc.footnotes && o.doc.footnotes.items.has(2), `seed ${seed}`);
    const all = ids(xml, await written(o, '/word/footnotes.xml'));
    assert.equal(new Set(all).size, all.length, `seed ${seed}: ${all.join(',')}`);
    while (o.history.canUndo) o.history.undo();
    assert.equal(docPicture(o.doc), picture, `seed ${seed}`);
    assert.equal(await written(o), original, `seed ${seed}`);
  }
});
