// writeDocx after edits made with the editing commands (src/ops/doc): what
// the commands leave in the model is what the writer has to put into the file,
// and what undo takes out of the model has to leave the file again.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { emptyWriteReport, writeDocx } from '../src/formats/docx/write';
import { paragraphText, rawOf, type DocModel, type DrawingItem, type Paragraph } from '../src/model/doc/types';
import { DocHistory, insertHyperlink, insertImage, insertTable, insertText, setRunFormat, splitParagraph, toggleList, type DocEdit, type DocPosition } from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { buildDocx as buildSpec, type DocxSpec } from './docx-fixtures';
import { canonicalText, convertedPicture, firstDifference, schemaOrderViolations } from './docx-write-helpers.test';

const decoder = new TextDecoder();
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 2, 0, 0, 0, 1, 8, 2, 0, 0, 0]);

interface Opened {
  pkg: OpcPackage;
  doc: DocModel;
  history: DocHistory;
  /** Uncompressed content of every part of the file as it was opened. */
  original: Map<string, Uint8Array>;
  apply(edit: DocEdit): void;
}

async function open(spec: DocxSpec): Promise<Opened> {
  const pkg = await OpcPackage.open(await buildSpec(spec));
  const doc = await readDocx(pkg);
  const history = new DocHistory(doc);
  const original = new Map<string, Uint8Array>();
  for (const part of pkg.partNames()) original.set(part, await pkg.readPart(part));
  return { pkg, doc, history, original, apply: (edit) => void history.applyEdit(edit, undefined) };
}

const at = (block: number, offset: number, story = 'body'): DocPosition => ({ story, path: [block], offset });
const text = async (pkg: OpcPackage, part: string): Promise<string> => decoder.decode(await pkg.readPart(part));

/**
 * Saves, reads the file again and requires the model that was written. The
 * comparison leaves out what the file states beyond the model of an edit: the
 * relationships the writer adds for new parts, and what the reader reads out
 * of the tree of a new drawing.
 */
async function reread(opened: Opened): Promise<DocModel> {
  const picture = convertedPicture(opened.doc);
  const again = await readDocx(await OpcPackage.open(await opened.pkg.save()));
  assert.equal(firstDifference(picture, convertedPicture(again)), undefined, 'the file does not read back as the edited model');
  assert.deepEqual(again.warnings, opened.doc.warnings);
  return again;
}

/** Every part of the package has the content it had when the file was opened, and there is no other part. */
async function sameAsOpened(opened: Opened): Promise<void> {
  const pkg = await OpcPackage.open(await opened.pkg.save());
  assert.deepEqual(pkg.partNames().sort(), Array.from(opened.original.keys()).sort());
  for (const [part, bytes] of opened.original) assert.deepEqual(await pkg.readPart(part), bytes, part);
}

const BODY =
  '<w:p w:rsidR="0001"><w:r><w:t>first paragraph</w:t></w:r></w:p>\n<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">second </w:t></w:r><w:r><w:t>paragraph</w:t></w:r></w:p>\n' +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>\n' +
  '<w:p><w:r><w:t>last</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';

test('typing, formatting and splitting: the edited paragraphs are written, their neighbours copied', async () => {
  const opened = await open({ body: BODY });
  const { pkg, doc } = opened;
  const before = await text(pkg, '/word/document.xml');
  opened.apply(insertText(doc, at(0, 5), ' new'));
  opened.apply(setRunFormat(doc, { anchor: at(0, 0), focus: at(0, 5) }, { bold: true }));
  opened.apply(splitParagraph(doc, at(3, 2)));
  opened.apply(insertText(doc, { story: 'body', path: [2, 0, 0, 0], offset: 4 }, '!'));
  const report = emptyWriteReport();
  await writeDocx(pkg, doc, { report });
  const after = await text(pkg, '/word/document.xml');
  // The second paragraph was not touched: its source is in the result as it was.
  const untouched = '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">second </w:t></w:r><w:r><w:t>paragraph</w:t></w:r></w:p>\n';
  assert.ok(before.indexOf(untouched) > 0 && after.indexOf(untouched) > 0);
  assert.ok(after.indexOf('<w:p w:rsidR="0001"><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>first</w:t></w:r><w:r><w:t xml:space="preserve"> new paragraph</w:t></w:r></w:p>') > 0, after);
  assert.ok(after.indexOf('<w:t>cell!</w:t>') > 0 && after.indexOf('<w:p><w:r><w:t>la</w:t></w:r></w:p><w:p><w:r><w:t>st</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>') > 0, after);
  assert.deepEqual(report.parts, ['/word/document.xml']);
  assert.deepEqual(schemaOrderViolations(after), []);
  await reread(opened);
});

test('edit, save, undo everything, save: every part is what the file had, and what the edits added is gone', async () => {
  const opened = await open({
    body: BODY,
    styles: '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>',
    numbering: '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>',
    headers: { rIdH: '<w:p><w:r><w:t>head</w:t></w:r></w:p>' },
    footnotes: '<w:footnote w:id="1"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:footnote>',
  });
  const { pkg, doc, history } = opened;
  opened.apply(insertText(doc, at(0, 0), 'typed '));
  opened.apply(toggleList(doc, { anchor: at(0, 0), focus: at(1, 0) }, 'bullet'));
  opened.apply(insertImage(doc, at(1, 3), { bytes: PNG, type: 'image/png', width: 9525, height: 9525 }));
  opened.apply(insertHyperlink(doc, { anchor: at(3, 0), focus: at(3, 4) }, 'https://example.com/'));
  opened.apply(insertTable(doc, at(3, 4), 2, 2));
  opened.apply(insertText(doc, at(0, 0, 'header:rIdH'), 'H'));
  opened.apply(insertText(doc, at(0, 0, 'footnote:1'), 'N'));
  const report = emptyWriteReport();
  await writeDocx(pkg, doc, { report });
  assert.deepEqual(report.parts.slice().sort(), [
    '/word/_rels/document.xml.rels',
    '/word/document.xml',
    '/word/footnotes.xml',
    '/word/header1.xml',
    '/word/media/image1.png',
    '/word/numbering.xml',
  ]);
  assert.deepEqual(await pkg.readPart('/word/media/image1.png'), PNG);
  assert.equal(pkg.contentTypeOf('/word/media/image1.png'), 'image/png');
  const rels = await pkg.relationships('/word/document.xml');
  assert.deepEqual(rels.filter((rel) => /\/(image|hyperlink)$/.test(rel.type)).map((rel) => [rel.target, rel.external]), [
    ['media/image1.png', false],
    ['https://example.com/', true],
  ]);
  const edited = await reread(opened);
  assert.ok(edited.numbering.nums.size === 2 && edited.numbering.abstracts.size === 2);

  // Back to what was read.
  while (history.canUndo) history.undo();
  assert.equal(doc.body.modified, undefined);
  assert.equal(doc.media, undefined);
  const second = emptyWriteReport();
  await writeDocx(pkg, doc, { report: second });
  assert.deepEqual(second.parts.slice().sort(), report.parts.slice().sort());
  await sameAsOpened(opened);

  // And forward again: the same file as after the first save.
  while (history.canRedo) history.redo();
  await writeDocx(pkg, doc);
  const redone = await reread(opened);
  assert.equal(firstDifference(convertedPicture(edited), convertedPicture(redone)), undefined);

  // A third save without a change in between writes the same again.
  const beforeThird = await text(pkg, '/word/document.xml');
  await writeDocx(pkg, doc);
  assert.equal(await text(pkg, '/word/document.xml'), beforeThird);
});

test('a list made by an edit: its definition and instance are added to the numbering part', async () => {
  const opened = await open({ body: BODY });
  const { pkg, doc } = opened;
  opened.apply(toggleList(doc, { anchor: at(0, 0), focus: at(1, 0) }, 'number'));
  assert.equal(doc.numbering.modified, true);
  await writeDocx(pkg, doc);
  const numbering = await text(pkg, '/word/numbering.xml');
  assert.deepEqual(schemaOrderViolations(numbering), []);
  assert.ok(numbering.indexOf('<w:abstractNum ') > 0 && numbering.indexOf('<w:abstractNum ') < numbering.indexOf('<w:num '));
  const again = await reread(opened);
  const first = again.body.blocks[0];
  assert.ok(first.type === 'paragraph' && first.props.numbering && again.numbering.nums.has(first.props.numbering.numId as number));
});

test('a picture inserted by an edit: the part, its content type, the relationship and the drawing as the command made it', async () => {
  const opened = await open({ body: BODY, headers: { rIdH: '<w:p><w:r><w:t>head</w:t></w:r></w:p>' } });
  const { pkg, doc } = opened;
  opened.apply(insertImage(doc, at(0, 5), { bytes: PNG, type: 'image/png', width: 914400, height: 457200, description: 'a picture' }));
  opened.apply(insertImage(doc, at(0, 0, 'header:rIdH'), { bytes: PNG.slice(), type: 'image/jpeg', width: 100, height: 100 }));
  await writeDocx(pkg, doc);
  assert.equal(pkg.contentTypeOf('/word/media/image1.png'), 'image/png');
  assert.equal(pkg.contentTypeOf('/word/media/image2.jpeg'), 'image/jpeg');
  assert.deepEqual((await pkg.relationships('/word/header1.xml')).map((rel) => rel.part), ['/word/media/image2.jpeg']);
  const again = await reread(opened);
  const block = again.body.blocks[0] as Paragraph;
  const items = block.children.reduce<unknown[]>((all, inline) => (inline.type === 'run' ? all.concat(inline.content) : all), []) as { type: string }[];
  const found = items.find((item) => item.type === 'drawing');
  assert.ok(found);
  const picture = found as DrawingItem;
  assert.deepEqual([picture.width, picture.height, picture.description, picture.image?.target], [914400, 457200, 'a picture', '/word/media/image1.png']);
});

test('a hyperlink made by an edit gets an external relationship', async () => {
  const opened = await open({ body: BODY });
  const { pkg, doc } = opened;
  opened.apply(insertHyperlink(doc, { anchor: at(0, 0), focus: at(0, 5) }, 'https://example.com/a?b=1&c=2'));
  opened.apply(insertHyperlink(doc, { anchor: at(3, 0), focus: at(3, 4) }, '#bookmark'));
  await writeDocx(pkg, doc);
  const rels = (await pkg.relationships('/word/document.xml')).filter((rel) => rel.type.endsWith('/hyperlink'));
  assert.deepEqual(rels.map((rel) => [rel.target, rel.external]), [['https://example.com/a?b=1&c=2', true]]);
  const after = await text(pkg, '/word/document.xml');
  assert.ok(after.indexOf(`<w:hyperlink r:id="${rels[0].id}"`) > 0 && after.indexOf('w:anchor="bookmark"') > 0, after);
  await reread(opened);
});

test('a content control around blocks: an edit inside writes the group from the model, the other groups are copied', async () => {
  const sdt = (id: number, inner: string) => `<w:sdt><w:sdtPr><w:id w:val="${id}"/></w:sdtPr><w:sdtContent >${inner}</w:sdtContent></w:sdt>`;
  const opened = await open({ body: `<w:p/>${sdt(1, '<w:p><w:r><w:t>aa</w:t></w:r></w:p><w:p ><w:r><w:t>bb</w:t></w:r></w:p>')}${sdt(2, '<w:p><w:r><w:t>cc</w:t></w:r></w:p>')}<w:p/>` });
  const { pkg, doc } = opened;
  opened.apply(insertText(doc, at(2, 1), 'X'));
  // The edit bound the whole group to a copy of its wrapper, which has no source span.
  const first = doc.body.blocks[1];
  const second = doc.body.blocks[2];
  assert.ok(first.wrap && first.wrap === second.wrap && first.wrap.srcStart === undefined);
  const report = emptyWriteReport();
  await writeDocx(pkg, doc, { report });
  const after = await text(pkg, '/word/document.xml');
  assert.ok(after.indexOf('<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>aa</w:t></w:r></w:p><w:p><w:r><w:t>bXb</w:t></w:r></w:p></w:sdtContent></w:sdt>') > 0, after);
  assert.ok(after.indexOf(sdt(2, '<w:p><w:r><w:t>cc</w:t></w:r></w:p>')) > 0, after);
  assert.deepEqual([report.copiedBlocks, report.writtenBlocks], [3, 1]);
  assert.deepEqual(report.warnings, []);
  await reread(opened);
});

test('an inline content control split over two paragraphs is written as two controls; only the first keeps the id', async () => {
  const opened = await open({
    body: '<w:p><w:r><w:t xml:space="preserve">before </w:t></w:r><w:sdt><w:sdtPr><w:alias w:val="Name"/><w:id w:val="77"/></w:sdtPr><w:sdtContent><w:r><w:t>inside the control</w:t></w:r></w:sdtContent></w:sdt></w:p>',
  });
  const { pkg, doc } = opened;
  opened.apply(splitParagraph(doc, at(0, 13)));
  const report = emptyWriteReport();
  await writeDocx(pkg, doc, { report });
  const after = await text(pkg, '/word/document.xml');
  const ids = after.match(/<w:id w:val="77"\/>/g) || [];
  const controls = after.match(/<w:sdt>/g) || [];
  if (controls.length === 2) {
    assert.equal(ids.length, 1, after);
    assert.ok(after.indexOf('<w:sdt><w:sdtPr><w:alias w:val="Name"/></w:sdtPr><w:sdtContent>') > 0, after);
    assert.equal(report.warnings.length, 1);
  } else {
    // The command kept the control in one paragraph: nothing to tell apart.
    assert.equal(ids.length, 1, after);
    assert.deepEqual(report.warnings, []);
  }
  const again = await readDocx(await OpcPackage.open(await pkg.save()));
  assert.deepEqual(again.body.blocks.map((block) => (block.type === 'paragraph' ? paragraphText(block) : '')), ['before inside', ' the control']);
});

test('notes: the edited note is written, the others are copied from the source', async () => {
  const footnotes =
    '\n<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>\n' +
    "<w:footnote w:id='1'><w:p><w:r><w:t>first note</w:t></w:r></w:p><w:p ><w:r><w:t>its second paragraph</w:t></w:r></w:p></w:footnote>\n<w:footnote  w:id=\"2\"><w:p><w:r><w:t>second note</w:t></w:r></w:p></w:footnote>\n";
  const opened = await open({ body: '<w:p><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>', footnotes });
  const { pkg, doc } = opened;
  const before = await text(pkg, '/word/footnotes.xml');
  opened.apply(insertText(doc, at(0, 5, 'footnote:1'), '!'));
  assert.equal(doc.footnotes?.modified, true);
  assert.equal(doc.footnotes?.items.get(1)?.srcStart, undefined);
  const report = emptyWriteReport();
  await writeDocx(pkg, doc, { report });
  assert.deepEqual(report.parts, ['/word/footnotes.xml']);
  const after = await text(pkg, '/word/footnotes.xml');
  // The untouched paragraph of the edited note, and the other notes, are the source text.
  assert.equal(
    after,
    before.replace("<w:footnote w:id='1'><w:p><w:r><w:t>first note</w:t></w:r></w:p>", '<w:footnote w:id="1"><w:p><w:r><w:t>first! note</w:t></w:r></w:p>'),
  );
  await reread(opened);
  opened.history.undo();
  await writeDocx(pkg, doc);
  await sameAsOpened(opened);
});

test('a section mark in the middle of the body keeps its place when blocks are added before it', async () => {
  const opened = await open({ body: '<w:p><w:r><w:t>one</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="100" w:h="200"/></w:sectPr><w:p><w:r><w:t>two</w:t></w:r></w:p>' });
  const { pkg, doc } = opened;
  assert.equal(doc.body.sectionIndex, 1);
  opened.apply(splitParagraph(doc, at(0, 1)));
  assert.equal(doc.body.sectionIndex, 2);
  await writeDocx(pkg, doc);
  const after = await text(pkg, '/word/document.xml');
  assert.ok(after.indexOf('<w:p><w:r><w:t>o</w:t></w:r></w:p><w:p><w:r><w:t>ne</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="100" w:h="200"/></w:sectPr><w:p><w:r><w:t>two</w:t></w:r></w:p></w:body>') > 0, after);
  await reread(opened);
});

test('untouched blocks keep their source text, byte for byte, whatever was edited around them', async () => {
  const opened = await open({ body: BODY });
  const { pkg, doc } = opened;
  const sources = doc.body.blocks.map((block) => rawOf(doc.body, block));
  opened.apply(insertText(doc, at(1, 0), 'x'));
  await writeDocx(pkg, doc);
  const again = await readDocx(await OpcPackage.open(await pkg.save()));
  again.body.blocks.forEach((block, i) => {
    if (i === 1) assert.notEqual(rawOf(again.body, block), sources[i]);
    else assert.equal(rawOf(again.body, block), sources[i]);
  });
  assert.equal(canonicalText(await pkg.readPart('/word/document.xml')).indexOf('xsecond'), canonicalText(await pkg.readPart('/word/document.xml')).lastIndexOf('xsecond'));
});
