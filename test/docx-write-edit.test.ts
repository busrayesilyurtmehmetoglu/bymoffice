// writeDocx on an edited document: only what was edited is written again, the
// rest of the part is copied from its source and other parts are not touched.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { compareIds, emptyWriteReport, writeDocx } from '../src/formats/docx/write';
import type { Block, DocModel, Paragraph } from '../src/model/doc/types';
import { OpcPackage } from '../src/opc';
import { ZipArchive } from '../src/zip';
import { buildDocx as buildSpec, documentXml, type DocxSpec } from './docx-fixtures';
import { schemaOrderViolations } from './docx-write-helpers.test';

const decoder = new TextDecoder();
const encoder = new TextEncoder();

async function open(spec: DocxSpec): Promise<{ pkg: OpcPackage; doc: DocModel; bytes: Uint8Array }> {
  const bytes = await buildSpec(spec);
  const pkg = await OpcPackage.open(bytes);
  return { pkg, doc: await readDocx(pkg), bytes };
}

const text = async (pkg: OpcPackage, part: string): Promise<string> => decoder.decode(await pkg.readPart(part));

const para = (value: string, bold = false): Paragraph => ({
  type: 'paragraph',
  props: {},
  children: [{ type: 'run', props: bold ? { bold: true } : {}, content: [{ type: 'text', text: value }] }],
});

/** Compressed bytes of every entry but the named ones: equal before and after means the entries were not rewritten. */
function untouched(before: Uint8Array, after: Uint8Array, except: string[]): void {
  const a = ZipArchive.read(before);
  const b = ZipArchive.read(after);
  assert.deepEqual(b.names().sort(), a.names().sort());
  for (const name of a.names()) {
    if (except.indexOf(name) >= 0) continue;
    const x = a.entry(name);
    const y = b.entry(name);
    assert.ok(x && y, name);
    assert.equal(x.data, undefined, name);
    assert.equal(y.data, undefined, name);
    assert.deepEqual(y.raw, x.raw, name);
  }
}

// Written with white space, attribute quotes and empty-element forms the writer itself would not produce.
const PRETTY_BODY =
  "\n    <w:p  w:rsidR='0001'>\n      <w:r><w:t>one</w:t></w:r>\n    </w:p>\n    <!-- between -->\n    <w:p><w:pPr><w:jc w:val = \"center\" ></w:jc></w:pPr><w:r><w:t>two</w:t></w:r></w:p>\n" +
  '    <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p></w:p></w:tc></w:tr></w:tbl>\n    <w:p><w:r><w:t>three</w:t></w:r></w:p>\n' +
  '    <w:sectPr><w:pgSz w:h="200" w:w="100"></w:pgSz></w:sectPr>\n  ';

test('nothing modified: the package is left as it is', async () => {
  const { pkg, doc, bytes } = await open({ body: PRETTY_BODY, styles: '', headers: { rIdH: '<w:p/>' } });
  const report = emptyWriteReport();
  await writeDocx(pkg, doc, { report });
  assert.deepEqual(report.parts, []);
  untouched(bytes, await pkg.save(), []);
});

test('a story marked as modified without an edit is written as the text it was read from', async () => {
  const { pkg, doc } = await open({ body: PRETTY_BODY, headers: { rIdH: '\n<w:p/>\n<!--x-->\n' }, footnotes: '\n <w:footnote w:id="1">\n  <w:p/>\n </w:footnote>\n' });
  const before = [await text(pkg, '/word/document.xml'), await text(pkg, '/word/header1.xml'), await text(pkg, '/word/footnotes.xml')];
  doc.body.modified = true;
  for (const story of doc.headers.values()) story.modified = true;
  if (doc.footnotes) doc.footnotes.modified = true;
  const report = emptyWriteReport();
  await writeDocx(pkg, doc, { report });
  assert.deepEqual([await text(pkg, '/word/document.xml'), await text(pkg, '/word/header1.xml'), await text(pkg, '/word/footnotes.xml')], before);
  assert.equal(report.writtenBlocks, 0);
  assert.equal(report.copiedBlocks, 5 + 2 + 1);
});

test('one edited paragraph: every other block, and the white space around them, is copied from the source', async () => {
  const { pkg, doc, bytes } = await open({ body: PRETTY_BODY, styles: '<w:style w:type="paragraph" w:styleId="A"/>' });
  const before = await text(pkg, '/word/document.xml');
  assert.equal(doc.body.blocks[2].type, 'paragraph');
  doc.body.blocks[2] = para('TWO', true);
  doc.body.modified = true;
  const report = emptyWriteReport();
  await writeDocx(pkg, doc, { report });
  const after = await text(pkg, '/word/document.xml');
  const old = '<w:p><w:pPr><w:jc w:val = "center" ></w:jc></w:pPr><w:r><w:t>two</w:t></w:r></w:p>';
  assert.ok(before.indexOf(old) > 0);
  assert.equal(after, before.replace(old, '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>TWO</w:t></w:r></w:p>'));
  assert.deepEqual(report.parts, ['/word/document.xml']);
  assert.deepEqual([report.copiedBlocks, report.writtenBlocks], [4, 1]);
  untouched(bytes, await pkg.save(), ['word/document.xml']);
  // And the edit is there when the file is read again.
  const again = await readDocx(await OpcPackage.open(await pkg.save()));
  const block = again.body.blocks[2];
  assert.ok(block.type === 'paragraph' && block.children[0].type === 'run');
  assert.deepEqual(block.children[0].content, [{ type: 'text', text: 'TWO' }]);
});

test('blocks added at the start and the end, and blocks removed', async () => {
  const { pkg, doc } = await open({ body: PRETTY_BODY });
  const blocks = doc.body.blocks;
  // Remove the comment and the table, add a paragraph before everything and one after.
  doc.body.blocks = [para('first'), blocks[0], blocks[2], blocks[4], para('last')];
  doc.body.modified = true;
  await writeDocx(pkg, doc);
  const after = await text(pkg, '/word/document.xml');
  const body = after.slice(after.indexOf('<w:body>') + 8, after.indexOf('</w:body>'));
  // The white space of the source stays where it was between what is copied; a new block stands after the space
  // that followed what was copied before it.
  assert.equal(
    body,
    "\n    <w:p><w:r><w:t>first</w:t></w:r></w:p><w:p  w:rsidR='0001'>\n      <w:r><w:t>one</w:t></w:r>\n    </w:p>" +
      '\n    <w:p><w:pPr><w:jc w:val = "center" ></w:jc></w:pPr><w:r><w:t>two</w:t></w:r></w:p>\n    <w:p><w:r><w:t>three</w:t></w:r></w:p>\n    ' +
      '<w:p><w:r><w:t>last</w:t></w:r></w:p><w:sectPr><w:pgSz w:h="200" w:w="100"></w:pgSz></w:sectPr>\n  ',
  );
  const again = await readDocx(await OpcPackage.open(await pkg.save()));
  assert.equal(again.body.blocks.length, 5);
});

test('the section mark at the end is copied while its span stands and written from the model otherwise', async () => {
  const { pkg, doc } = await open({ body: PRETTY_BODY });
  doc.body.section = { page: { width: 300, height: 400 }, titlePage: true };
  delete doc.body.sectionSpan;
  doc.body.modified = true;
  await writeDocx(pkg, doc);
  const after = await text(pkg, '/word/document.xml');
  assert.ok(after.indexOf('<w:sectPr><w:pgSz w:w="300" w:h="400"/><w:titlePg/></w:sectPr>') > 0, after);
  assert.ok(after.indexOf("<w:p  w:rsidR='0001'>") > 0);
  assert.deepEqual(schemaOrderViolations(after), []);
});

test('a wrapper around blocks is copied whole while its content is untouched, and written from the model after an edit in it', async () => {
  const sdt = (id: number, inner: string) => `<w:sdt><w:sdtPr><w:id w:val="${id}"/></w:sdtPr><w:sdtContent >${inner}</w:sdtContent></w:sdt>`;
  const body = `<w:p/>${sdt(1, '<w:p><w:r><w:t>a</w:t></w:r></w:p><w:p ><w:r><w:t>b</w:t></w:r></w:p>')}${sdt(2, '<w:p><w:r><w:t>c</w:t></w:r></w:p>')}<w:p/>`;
  const { pkg, doc } = await open({ body });
  // The blocks in a wrapper carry no span of their own; the wrapper does.
  const inner = doc.body.blocks[2];
  assert.ok(inner.type === 'paragraph' && inner.wrap && inner.srcStart === undefined && inner.wrap.srcStart !== undefined);
  // The edit keeps the wrapper object, as an edit of text in a content control would.
  doc.body.blocks[2] = { ...para('B'), wrap: inner.wrap };
  doc.body.modified = true;
  const report = emptyWriteReport();
  await writeDocx(pkg, doc, { report });
  const after = await text(pkg, '/word/document.xml');
  assert.ok(after.indexOf(`<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>a</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p></w:sdtContent></w:sdt>`) > 0, after);
  // The other one keeps its source, with the space in its tag.
  assert.ok(after.indexOf(sdt(2, '<w:p><w:r><w:t>c</w:t></w:r></w:p>')) > 0, after);
  assert.deepEqual([report.copiedBlocks, report.writtenBlocks], [3, 1]);
});

test('only the part of the modified story is written: a header', async () => {
  const { pkg, doc, bytes } = await open({
    body: '<w:p/><w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/></w:sectPr>',
    headers: { rIdH: '<w:p><w:r><w:t>head</w:t></w:r></w:p><w:p/>' },
    footers: { rIdF: '<w:p><w:r><w:t>foot</w:t></w:r></w:p>' },
  });
  const header = doc.headers.get('rIdH');
  assert.ok(header);
  header.blocks[0] = para('HEAD');
  header.modified = true;
  const report = emptyWriteReport();
  await writeDocx(pkg, doc, { report });
  assert.deepEqual(report.parts, ['/word/header1.xml']);
  untouched(bytes, await pkg.save(), ['word/header1.xml']);
  const after = await text(pkg, '/word/header1.xml');
  assert.ok(after.indexOf('<w:p><w:r><w:t>HEAD</w:t></w:r></w:p><w:p/></w:hdr>') > 0, after);
});

test('notes: the edited block is written, the other notes and the tags of the edited one are copied', async () => {
  const footnotes =
    '\n<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>\n' +
    "<w:footnote w:id='1'  w:custom=\"x\">\n <w:p><w:r><w:t>first</w:t></w:r></w:p>\n <w:p><w:r><w:t>second</w:t></w:r></w:p>\n</w:footnote >\n<w:footnote w:id=\"2\"/>\n";
  const { pkg, doc } = await open({ body: '<w:p/>', footnotes });
  const before = await text(pkg, '/word/footnotes.xml');
  const notes = doc.footnotes;
  assert.ok(notes);
  const note = notes.items.get(1);
  assert.ok(note);
  note.blocks[1] = para('SECOND');
  notes.modified = true;
  await writeDocx(pkg, doc);
  const after = await text(pkg, '/word/footnotes.xml');
  assert.equal(after, before.replace('<w:p><w:r><w:t>second</w:t></w:r></w:p>\n</w:footnote >', '<w:p><w:r><w:t>SECOND</w:t></w:r></w:p></w:footnote>'));

  // A note that is new in the model is written from it; an empty one gets content.
  notes.items.set(7, { id: 7, blocks: [para('new note')] });
  const empty = notes.items.get(2);
  assert.ok(empty);
  empty.blocks.push(para('now with text'));
  await writeDocx(pkg, doc);
  const third = await text(pkg, '/word/footnotes.xml');
  assert.ok(third.indexOf('<w:footnote w:id="2"><w:p><w:r><w:t>now with text</w:t></w:r></w:p></w:footnote>') > 0, third);
  assert.ok(third.indexOf('<w:footnote w:id="7"><w:p><w:r><w:t>new note</w:t></w:r></w:p></w:footnote></w:footnotes>') > 0, third);
  const again = await readDocx(await OpcPackage.open(await pkg.save()));
  assert.deepEqual(Array.from(again.footnotes?.items.keys() || []), [-1, 1, 2, 7]);
});

test('a relationship the model has and the part lacks is added; none is removed', async () => {
  const { pkg, doc } = await open({
    body: '<w:p><w:hyperlink r:id="rIdL1"><w:r><w:t>old</w:t></w:r></w:hyperlink></w:p>',
    rels: [
      ['rIdL1', 'hyperlink', 'https://example.com/old', true],
      ['rIdUnused', 'hyperlink', 'https://example.com/unused', true],
    ],
  });
  const type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
  doc.body.relationships.set('rIdNew', { id: 'rIdNew', type, target: 'https://example.com/a?b=1&c="2"', external: true, part: null });
  doc.body.blocks.push({ type: 'paragraph', props: {}, children: [{ type: 'run', props: {}, content: [{ type: 'text', text: 'new' }], wrap: { kind: 'hyperlink', rId: 'rIdNew', tooltip: 'a "tip" <b>' } }] });
  doc.body.modified = true;
  await writeDocx(pkg, doc);
  const rels = await pkg.relationships('/word/document.xml');
  // (A part that is written anew lists its relationships by id: see writeDocx.)
  assert.deepEqual(
    rels.map((rel) => rel.id),
    ['rIdL1', 'rIdNew', 'rIdUnused'],
  );
  assert.deepEqual(rels[1], { id: 'rIdNew', type, target: 'https://example.com/a?b=1&c="2"', external: true, part: null });
  const again = await readDocx(await OpcPackage.open(await pkg.save()));
  const added = again.body.blocks[1];
  assert.ok(added.type === 'paragraph' && added.children[0].wrap && added.children[0].wrap.kind === 'hyperlink');
  assert.equal(added.children[0].wrap.target, 'https://example.com/a?b=1&c="2"');
  assert.equal(added.children[0].wrap.tooltip, 'a "tip" <b>');
});

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 2, 0, 0, 0, 1, 8, 2, 0, 0, 0]);

test('a new picture: its part, its content type, its relationship and a standard drawing', async () => {
  const { pkg, doc } = await open({ body: '<w:p/>', headers: { rIdH: '<w:p/>' } });
  doc.media = new Map([['/word/media/new1.png', { bytes: PNG, type: 'image/png' }]]);
  const picture = (): Block => ({
    type: 'paragraph',
    props: {},
    children: [
      {
        type: 'run',
        props: {},
        content: [
          {
            type: 'drawing',
            placement: 'inline',
            kind: 'picture',
            width: 914400,
            height: 457200,
            description: 'a "new" picture',
            image: { rId: '', target: '/word/media/new1.png' },
            xml: { type: 'element', name: 'none', attrs: new Map(), children: [] },
          },
        ],
      },
    ],
  });
  doc.body.blocks.push(picture(), picture());
  doc.body.modified = true;
  const header = doc.headers.get('rIdH');
  assert.ok(header);
  header.blocks.push(picture());
  header.modified = true;
  await writeDocx(pkg, doc);
  assert.deepEqual(await pkg.readPart('/word/media/new1.png'), PNG);
  assert.equal(pkg.contentTypeOf('/word/media/new1.png'), 'image/png');
  // One relationship for the two drawings of the body, one of its own for the header.
  const rels = (await pkg.relationships('/word/document.xml')).filter((rel) => rel.type.endsWith('/image'));
  assert.deepEqual(rels.map((rel) => [rel.target, rel.part]), [['media/new1.png', '/word/media/new1.png']]);
  const headerRels = await pkg.relationships('/word/header1.xml');
  assert.deepEqual(headerRels.map((rel) => rel.part), ['/word/media/new1.png']);

  const again = await readDocx(await OpcPackage.open(await pkg.save()));
  for (const block of [again.body.blocks[1], again.body.blocks[2], again.headers.get('rIdH')?.blocks[1]]) {
    assert.ok(block && block.type === 'paragraph' && block.children[0].type === 'run');
    const item = block.children[0].content[0];
    assert.ok(item.type === 'drawing');
    assert.deepEqual([item.placement, item.kind, item.width, item.height, item.description, item.image?.target], ['inline', 'picture', 914400, 457200, 'a "new" picture', '/word/media/new1.png']);
  }
  assert.deepEqual(again.warnings, []);
});

const NUMBERING =
  '\n<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>\n' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>\n<w:numIdMacAtCleanup w:val="1"/>\n';

function addList(doc: DocModel): void {
  doc.numbering.abstracts.set(5, {
    id: 5,
    multiLevelType: 'hybridMultilevel',
    levels: new Map([
      [0, { level: 0, start: 1, format: 'bullet', text: '•', align: 'left', suffix: 'tab', para: { indent: { left: 720, hanging: 360 } }, run: { fonts: { ascii: 'Symbol', hAnsi: 'Symbol' } } }],
      [1, { level: 1, start: 3, format: 'lowerLetter', text: '%2)', align: 'left', restart: 0, legal: true, paraStyle: 'ListParagraph' }],
    ]),
  });
  doc.numbering.nums.set(9, { id: 9, abstractId: 5, overrides: new Map([[1, { start: 4 }]]) });
  doc.numbering.nums.set(10, { id: 10, abstractId: 0, overrides: new Map([[0, { level: { level: 0, start: 7, format: 'upperRoman', text: '%1' } }]]) });
  doc.numbering.modified = true;
}

test('numbering definitions the model gained are added to the text of the numbering part', async () => {
  const { pkg, doc } = await open({ body: '<w:p/>', numbering: NUMBERING });
  const before = await text(pkg, '/word/numbering.xml');
  addList(doc);
  const report = emptyWriteReport();
  await writeDocx(pkg, doc, { report });
  assert.deepEqual(report.parts, ['/word/numbering.xml']);
  const after = await text(pkg, '/word/numbering.xml');
  const abstract =
    '<w:abstractNum w:abstractNumId="5"><w:multiLevelType w:val="hybridMultilevel"/>' +
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:suff w:val="tab"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl>' +
    '<w:lvl w:ilvl="1"><w:start w:val="3"/><w:numFmt w:val="lowerLetter"/><w:lvlRestart w:val="0"/><w:pStyle w:val="ListParagraph"/><w:isLgl/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>';
  const nums =
    '<w:num w:numId="9"><w:abstractNumId w:val="5"/><w:lvlOverride w:ilvl="1"><w:startOverride w:val="4"/></w:lvlOverride></w:num>' +
    '<w:num w:numId="10"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0"><w:start w:val="7"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1"/></w:lvl></w:lvlOverride></w:num>';
  // Abstract definitions before the first instance, instances after the last, the clean-up mark still last; the rest as it was.
  assert.equal(after, before.replace('<w:num w:numId="1">', abstract + '<w:num w:numId="1">').replace('<w:numIdMacAtCleanup', nums + '<w:numIdMacAtCleanup'));
  assert.deepEqual(schemaOrderViolations(after), []);
  const again = await readDocx(await OpcPackage.open(await pkg.save()));
  assert.deepEqual(again.numbering, { abstracts: doc.numbering.abstracts, nums: doc.numbering.nums });

  // Written a second time, nothing is added twice.
  await writeDocx(pkg, doc);
  assert.equal(await text(pkg, '/word/numbering.xml'), after);
});

test('a numbering part is made when the file has none', async () => {
  const { pkg, doc } = await open({ body: '<w:p/>' });
  addList(doc);
  doc.numbering.abstracts.set(0, { id: 0, levels: new Map() });
  await writeDocx(pkg, doc);
  assert.equal(pkg.contentTypeOf('/word/numbering.xml'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml');
  const rels = (await pkg.relationships('/word/document.xml')).filter((rel) => rel.type.endsWith('/numbering'));
  assert.deepEqual(rels.map((rel) => rel.part), ['/word/numbering.xml']);
  const again = await readDocx(await OpcPackage.open(await pkg.save()));
  assert.deepEqual(again.numbering, { abstracts: doc.numbering.abstracts, nums: doc.numbering.nums });
  assert.deepEqual(again.warnings, []);
});

test('a style the model gained is added to the styles part; the styles that were there are not rewritten', async () => {
  const styles =
    '\n<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val = "22"/></w:rPr></w:rPrDefault></w:docDefaults>\n<w:latentStyles w:count="1"/>\n' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>\n';
  const { pkg, doc } = await open({ body: '<w:p/>', styles });
  const before = await text(pkg, '/word/styles.xml');
  doc.styles.styles.set('MyStyle', {
    id: 'MyStyle',
    type: 'paragraph',
    isDefault: false,
    custom: true,
    name: 'My Style',
    basedOn: 'Normal',
    next: 'Normal',
    link: 'MyStyleChar',
    uiPriority: 9,
    quickFormat: true,
    semiHidden: false,
    para: { keepNext: true, spacing: { before: 240 } },
    run: { bold: true, size: 28 },
  });
  doc.styles.styles.set('MyTable', {
    id: 'MyTable',
    type: 'table',
    isDefault: false,
    name: 'My Table',
    table: { borders: { top: { style: 'single', size: 4 } } },
    row: { cantSplit: true },
    cell: { vAlign: 'center' },
    conditions: new Map([['firstRow', { run: { bold: true }, cell: { shading: { pattern: 'clear', fill: { rgb: 'EEEEEE' } } } }]]),
  });
  // Styles are looked at when a story is written: a style nothing uses need not be saved.
  doc.body.blocks.push({ type: 'paragraph', props: { style: 'MyStyle' }, children: [] });
  doc.body.modified = true;
  await writeDocx(pkg, doc);
  const after = await text(pkg, '/word/styles.xml');
  const added =
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="MyStyle"><w:name w:val="My Style"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:link w:val="MyStyleChar"/>' +
    '<w:uiPriority w:val="9"/><w:semiHidden w:val="0"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>' +
    '<w:style w:type="table" w:styleId="MyTable"><w:name w:val="My Table"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:trPr><w:cantSplit/></w:trPr><w:tcPr><w:vAlign w:val="center"/></w:tcPr>' +
    '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr><w:tcPr><w:shd w:val="clear" w:fill="EEEEEE"/></w:tcPr></w:tblStylePr></w:style>';
  assert.equal(after, before.replace('</w:styles>', added + '</w:styles>'));
  assert.deepEqual(schemaOrderViolations(after), []);
  const again = await readDocx(await OpcPackage.open(await pkg.save()));
  assert.deepEqual(again.styles.styles.get('MyStyle'), doc.styles.styles.get('MyStyle'));
  assert.deepEqual(again.styles.styles.get('MyTable'), doc.styles.styles.get('MyTable'));
});

test('a story that is no longer modified goes back to what the file had', async () => {
  const { pkg, doc, bytes } = await open({ body: PRETTY_BODY });
  const before = await text(pkg, '/word/document.xml');
  const original = doc.body.blocks[0];
  doc.body.blocks[0] = para('changed');
  doc.body.modified = true;
  await writeDocx(pkg, doc);
  assert.notEqual(await text(pkg, '/word/document.xml'), before);
  // Undone: the block is the original again and the story is clean.
  doc.body.blocks[0] = original;
  doc.body.modified = false;
  await writeDocx(pkg, doc);
  assert.equal(await text(pkg, '/word/document.xml'), before);
  assert.deepEqual(await (await OpcPackage.open(await pkg.save())).readPart('/word/document.xml'), await (await OpcPackage.open(bytes)).readPart('/word/document.xml'));
});

test('a second edit after a save is written against the source, not against the first result', async () => {
  const { pkg, doc } = await open({ body: PRETTY_BODY });
  doc.body.blocks[0] = para('first edit');
  doc.body.modified = true;
  await writeDocx(pkg, doc);
  doc.body.blocks[4] = para('second edit');
  await writeDocx(pkg, doc);
  const after = await text(pkg, '/word/document.xml');
  assert.ok(after.indexOf('first edit') > 0 && after.indexOf('second edit') > 0 && after.indexOf('<w:t>two</w:t>') > 0 && after.indexOf('<w:t>three</w:t>') < 0);
  const again = await readDocx(await OpcPackage.open(await pkg.save()));
  assert.equal(again.body.blocks.length, 5);
});

test('a changed root element is written from the model; the blocks are still copied', async () => {
  const { pkg, doc } = await open({ body: PRETTY_BODY });
  doc.body.rootAttrs = doc.body.rootAttrs.concat([['xmlns:new', 'urn:new']]);
  doc.body.modified = true;
  const report = emptyWriteReport();
  await writeDocx(pkg, doc, { report });
  const after = await text(pkg, '/word/document.xml');
  assert.ok(after.indexOf(' xmlns:new="urn:new"><w:body>\n    <w:p  w:rsidR=\'0001\'>') > 0, after.slice(0, 900));
  assert.equal(report.writtenBlocks, 0);
});

test('a part in UTF-16 stays in UTF-16', async () => {
  const source = documentXml('<w:p><w:r><w:t>ğüşiöç</w:t></w:r></w:p><w:p/>').replace('UTF-8', 'UTF-16');
  const bytes = new Uint8Array(2 + source.length * 2);
  bytes[0] = 0xff;
  bytes[1] = 0xfe;
  for (let i = 0; i < source.length; i++) {
    bytes[2 + i * 2] = source.charCodeAt(i) & 0xff;
    bytes[3 + i * 2] = source.charCodeAt(i) >> 8;
  }
  const { pkg, doc } = await open({ document: bytes });
  doc.body.blocks[1] = para('İkinci');
  doc.body.modified = true;
  await writeDocx(pkg, doc);
  const written = await pkg.readPart('/word/document.xml');
  assert.deepEqual([written[0], written[1]], [0xff, 0xfe]);
  const again = await readDocx(await OpcPackage.open(await pkg.save()));
  const block = again.body.blocks[1];
  assert.ok(block.type === 'paragraph' && block.children[0].type === 'run');
  assert.deepEqual(block.children[0].content, [{ type: 'text', text: 'İkinci' }]);
});

test('a model that was not read from a package cannot be written into one', async () => {
  const { pkg, doc } = await open({ body: '<w:p/>' });
  const loose: DocModel = { ...doc, body: { ...doc.body, part: 'doc:body', rootName: '', rootAttrs: [], source: '' } };
  await assert.rejects(writeDocx(pkg, loose), /buildDocx/);
  // Nothing was put into the package.
  assert.equal(encoder.encode(await text(pkg, '/word/document.xml')).length > 0, true);
});

test('the ids of relationships are put in one order, whatever order they come in', () => {
  const sets = [
    ['rId1000000000000000', 'rId9', 'rId10'],
    ['rId01', 'rId1', 'rId001', 'rId2'],
    ['rId99999999999999999999', 'rId100000000000000000000', 'rIdA', 'rId', 'rIdS1', 'rIdS01', 'rId0'],
  ];
  let seed = 7;
  const random = (): number => ((seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648);
  for (let i = 0; i < 40; i++) {
    const set = new Set<string>();
    while (set.size < 6) set.add('rId' + ['', 'S', 'C'][Math.floor(random() * 3)] + '0'.repeat(Math.floor(random() * 3)) + String(Math.floor(random() * 10 ** Math.floor(random() * 22))));
    sets.push(Array.from(set));
  }
  const permutations = (list: string[]): string[][] => (list.length <= 1 ? [list] : list.reduce((all: string[][], item, i) => all.concat(permutations(list.slice(0, i).concat(list.slice(i + 1))).map((rest) => [item].concat(rest))), []));
  for (const set of sets) {
    const orders = new Set<string>();
    for (const order of permutations(set.slice(0, 6))) orders.add(order.slice().sort(compareIds).join(' '));
    assert.equal(orders.size, 1, `${set.join(' ')}: ${Array.from(orders).join(' | ')}`);
  }
  assert.deepEqual(['rId10', 'rId9', 'rId1000000000000000000'].sort(compareIds), ['rId9', 'rId10', 'rId1000000000000000000']);
});

