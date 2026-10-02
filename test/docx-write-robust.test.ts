// The docx writer under stress: deep nesting, large documents, text that XML
// cannot hold, and models that are not what a reader would have made.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { buildDocx, emptyWriteReport, writeDocx } from '../src/formats/docx/write';
import { defaultSettings, emptyDocStyles, emptyTheme, type Block, type DocModel, type Paragraph, type Wrapper } from '../src/model/doc/types';
import { OpcPackage } from '../src/opc';
import { parseXml } from '../src/xml';
import { buildDocx as buildSpec } from './docx-fixtures';
import { documentPicture, firstDifference } from './docx-write-helpers.test';

const decoder = new TextDecoder();

const para = (text: string): Paragraph => ({ type: 'paragraph', props: {}, children: [{ type: 'run', props: {}, content: [{ type: 'text', text }] }] });

function model(blocks: Block[]): DocModel {
  return {
    body: { part: 'doc:body', rootName: '', rootAttrs: [], blocks, relationships: new Map(), source: '' },
    styles: emptyDocStyles(),
    numbering: { abstracts: new Map(), nums: new Map() },
    theme: emptyTheme(),
    settings: defaultSettings(),
    fonts: new Map(),
    headers: new Map(),
    footers: new Map(),
    strict: false,
    warnings: [],
    unknown: new Map(),
  };
}

test('tables, content controls and text boxes nested as deep as the reader allows are written without running out of stack', async () => {
  const depth = 60;
  let tables = '<w:p><w:r><w:t>innermost</w:t></w:r></w:p>';
  for (let i = 0; i < depth; i++) tables = `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc>${tables}<w:p/></w:tc></w:tr></w:tbl>`;
  let controls = '<w:p><w:r><w:t>innermost</w:t></w:r></w:p>';
  for (let i = 0; i < depth; i++) controls = `<w:sdt><w:sdtPr><w:id w:val="${i}"/></w:sdtPr><w:sdtContent>${controls}</w:sdtContent></w:sdt>`;
  let inline = '<w:r><w:t>innermost</w:t></w:r>';
  for (let i = 0; i < depth; i++) inline = i % 2 ? `<w:ins w:id="${i}" w:author="A">${inline}</w:ins>` : `<w:hyperlink w:anchor="a${i}">${inline}</w:hyperlink>`;
  const pkg = await OpcPackage.open(await buildSpec({ body: `${tables}${controls}<w:p>${inline}</w:p>` }));
  const doc = await readDocx(pkg);
  const picture = documentPicture(doc);
  await writeDocx(pkg, doc, { regenerate: true });
  const again = await readDocx(await OpcPackage.open(await pkg.save()));
  assert.equal(firstDifference(picture, documentPicture(again)), undefined);
});

test('a document of 50,000 paragraphs: one edit is written in a fraction of a second, a full rewrite in a few', async (t) => {
  const count = 50000;
  const parts: string[] = [];
  for (let i = 0; i < count; i++) parts.push(`<w:p w:rsidR="00A1B2C3"><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Paragraph ${i} of the document</w:t></w:r></w:p>`);
  const pkg = await OpcPackage.open(await buildSpec({ body: parts.join('') + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>' }));
  const doc = await readDocx(pkg);
  assert.equal(doc.body.blocks.length, count);
  doc.body.blocks[25000] = para('edited');
  doc.body.modified = true;
  const report = emptyWriteReport();
  let started = Date.now();
  await writeDocx(pkg, doc, { report });
  const one = Date.now() - started;
  assert.deepEqual([report.copiedBlocks, report.writtenBlocks], [count - 1, 1]);
  const written = decoder.decode(await pkg.readPart('/word/document.xml'));
  assert.ok(written.indexOf('<w:t>Paragraph 24999 of the document</w:t></w:r></w:p><w:p><w:r><w:t>edited</w:t></w:r></w:p><w:p w:rsidR="00A1B2C3">') > 0);
  assert.equal(written.length, doc.body.source.length - parts[25000].length + '<w:p><w:r><w:t>edited</w:t></w:r></w:p>'.length);

  started = Date.now();
  await writeDocx(pkg, doc, { regenerate: true });
  const all = Date.now() - started;
  t.diagnostic(`one edited paragraph of ${count}: ${one} ms; every block from the model: ${all} ms`);
  // The aim is under 300 ms for the edit; the bounds here leave room for a busy machine.
  assert.ok(one < 1500, `writing one edit took ${one} ms`);
  assert.ok(all < 15000, `writing everything took ${all} ms`);
  const again = await readDocx(await OpcPackage.open(await pkg.save()));
  assert.equal(again.body.blocks.length, count);
});

test('characters XML cannot hold are left out and counted; everything else is escaped', async () => {
  const text = 'a\u0000b\u0001c\u000bd\u001fe\ud800f\udc00g￾h￿i 😀 <&>"\' ]]> \t\n\r end';
  const source = model([
    {
      type: 'paragraph',
      props: { style: 'x"<&>\u0002' },
      attrs: [['w:rsidR', 'a"b<c&d\te\nf\u0003']],
      children: [
        { type: 'run', props: {}, content: [{ type: 'text', text }, { type: 'symbol', font: 'F"\u0004', char: '' }, { type: 'instrText', text: ' HYPERLINK "x&y" \u0005' }] },
        { type: 'bookmarkStart', id: '1"', name: 'b<\u0006>' },
        { type: 'misc', node: { type: 'comment', value: 'a -- b\u0007-' } },
        { type: 'misc', node: { type: 'pi', target: 'pi', value: 'x ?> y' } },
      ],
    },
  ]);
  const report = emptyWriteReport();
  const bytes = await buildDocx(source, { report });
  // 0000, 0001, 000B, 001F, two lone surrogates, FFFE, FFFF in the text; one each in the style, the attribute, the font, the field code, the name and the comment.
  assert.equal(report.droppedCharacters, 8 + 6);
  const pkg = await OpcPackage.open(bytes);
  const xml = decoder.decode(await pkg.readPart('/word/document.xml'));
  parseXml(xml);
  assert.ok(!/[\u0000-\u0008\u000b\u000c\u000e-\u001f￾￿]/.test(xml));
  const again = await readDocx(pkg);
  const paragraph = again.body.blocks[0];
  assert.ok(paragraph.type === 'paragraph' && paragraph.children[0].type === 'run');
  assert.deepEqual(paragraph.children[0].content[0], { type: 'text', text: 'abcdefghi 😀 <&>"\' ]]> \t\n\r end' });
  assert.equal(paragraph.props.style, 'x"<&>');
  assert.deepEqual(paragraph.attrs, [['w:rsidR', 'a"b<c&d\te\nf']]);
  assert.deepEqual(paragraph.children[1], { type: 'bookmarkStart', id: '1"', name: 'b<>' });
  assert.ok(xml.indexOf('<!--a - b- -->') > 0 && xml.indexOf('<?pi x ? > y?>') > 0, xml);
});

test('a kept tree with a character XML cannot hold is written without it', async () => {
  const pkg = await OpcPackage.open(await buildSpec({ body: '<w:p><w:r><w:t>a&#1;b</w:t><w:mystery w:x="c&#2;d">e&#3;f</w:mystery></w:r></w:p>' }));
  const doc = await readDocx(pkg);
  const report = emptyWriteReport();
  await writeDocx(pkg, doc, { regenerate: true, report });
  assert.equal(report.droppedCharacters, 3);
  const xml = decoder.decode(await pkg.readPart('/word/document.xml'));
  assert.ok(xml.indexOf('<w:t>ab</w:t><w:mystery w:x="cd">ef</w:mystery>') > 0, xml);
});

test('a model no reader would make is refused with an error, not written wrongly', async () => {
  // A chain of wrappers that never ends.
  const loop: Wrapper = { kind: 'ins' };
  loop.parent = loop;
  await assert.rejects(buildDocx(model([{ type: 'paragraph', props: {}, children: [{ type: 'run', props: {}, content: [], wrap: loop }] }])), /wrappers/);
  // Tables nested beyond any document.
  let block: Block = para('x');
  for (let i = 0; i < 400; i++) block = { type: 'table', props: {}, grid: [1], rows: [{ props: {}, cells: [{ props: {}, blocks: [block] }] }] };
  await assert.rejects(buildDocx(model([block])), /nested deeper/);
});

test('a failure while writing leaves the package as it was', async () => {
  const pkg = await OpcPackage.open(await buildSpec({ body: '<w:p><w:r><w:t>kept</w:t></w:r></w:p>', headers: { rIdH: '<w:p/>' } }));
  const doc = await readDocx(pkg);
  const before = decoder.decode(await pkg.readPart('/word/document.xml'));
  doc.body.blocks[0] = para('changed');
  doc.body.modified = true;
  const loop: Wrapper = { kind: 'ins' };
  loop.parent = loop;
  const header = doc.headers.get('rIdH');
  assert.ok(header);
  header.blocks[0] = { type: 'paragraph', props: {}, children: [{ type: 'run', props: {}, content: [], wrap: loop }] };
  header.modified = true;
  await assert.rejects(writeDocx(pkg, doc), /wrappers/);
  assert.equal(decoder.decode(await pkg.readPart('/word/document.xml')), before);
});
