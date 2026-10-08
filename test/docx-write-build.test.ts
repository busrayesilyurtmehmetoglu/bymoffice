// buildDocx: a new package from a model alone. The result is read back with
// the docx reader and compared with the model it was built from.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { NumberingCounters } from '../src/formats/docx/numbering';
import { buildDocx, emptyWriteReport } from '../src/formats/docx/write';
import {
  defaultSettings,
  emptyDocStyles,
  emptyTheme,
  paragraphText,
  type Block,
  type DocModel,
  type DrawingItem,
  type Paragraph,
  type Story,
  type StoryCollection,
  type Wrapper,
} from '../src/model/doc/types';
import { OpcPackage } from '../src/opc';
import { parseXml } from '../src/xml';
import { simpleDoc } from './doc-binary-fixtures';
import { buildDocx as buildSpec, THEME } from './docx-fixtures';
import { convertedPicture, firstDifference, schemaOrderViolations } from './docx-write-helpers.test';
import { doc as rtf, read as readRtfSource } from './rtf-fixtures';

const decoder = new TextDecoder();

const para = (text: string, props: Paragraph['props'] = {}): Paragraph => ({
  type: 'paragraph',
  props,
  children: text ? [{ type: 'run', props: {}, content: [{ type: 'text', text }] }] : [],
});

const story = (part: string, blocks: Block[]): Story => ({ part, rootName: '', rootAttrs: [], blocks, relationships: new Map(), source: '' });

function collection<T>(part: string, items: [number, T][]): StoryCollection<T> {
  return { part, rootName: '', rootAttrs: [], relationships: new Map(), source: '', items: new Map(items) };
}

function model(blocks: Block[], more: Partial<DocModel> = {}): DocModel {
  return {
    body: story('doc:body', blocks),
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
    ...more,
  };
}

interface Built {
  pkg: OpcPackage;
  result: DocModel;
  part(name: string): Promise<string>;
}

async function build(source: DocModel, options?: Parameters<typeof buildDocx>[1]): Promise<Built> {
  const bytes = await buildDocx(source, options);
  const pkg = await OpcPackage.open(bytes);
  const result = await readDocx(pkg);
  assert.deepEqual(result.warnings, []);
  // Every XML part parses, and the children of property elements are in the order of the schema.
  for (const name of pkg.partNames()) {
    if (!/\.(xml|rels)$/.test(name)) continue;
    const data = await pkg.readPart(name);
    parseXml(data);
    assert.deepEqual(schemaOrderViolations(data), [], name);
  }
  return { pkg, result, part: async (name) => decoder.decode(await pkg.readPart(name)) };
}

/** Builds, reads back and requires the same model, but for what a new package states differently. */
async function same(source: DocModel): Promise<Built> {
  const built = await build(source);
  assert.equal(firstDifference(convertedPicture(source), convertedPicture(built.result)), undefined);
  return built;
}

test('the package of a small model: parts, content types and relationships', async () => {
  const source = model([para('Merhaba dünya')]);
  source.body.section = { page: { width: 11906, height: 16838 }, margins: { top: 1417, right: 1417, bottom: 1417, left: 1417 } };
  const built = await same(source);
  assert.deepEqual(built.pkg.partNames().sort(), ['/_rels/.rels', '/word/_rels/document.xml.rels', '/word/document.xml', '/word/settings.xml', '/word/styles.xml']);
  assert.equal(built.pkg.contentTypeOf('/word/document.xml'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml');
  assert.equal(await built.pkg.mainPart(), '/word/document.xml');
  const document = await built.part('/word/document.xml');
  assert.ok(document.indexOf('<w:body><w:p><w:r><w:t>Merhaba dünya</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417"/></w:sectPr></w:body></w:document>') > 0, document);
  assert.ok(document.indexOf('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>') === 0);
  const settings = await built.part('/word/settings.xml');
  assert.ok(settings.indexOf('<w:defaultTabStop w:val="720"/></w:settings>') > 0, settings);
});

test('a document read from RTF comes back the same from the docx built of it', async () => {
  const source = readRtfSource(
    rtf(
      '{\\header head \\chpgn\\par}{\\footer foot\\par}' +
        '\\pard\\qc\\sb120\\sa240\\li720\\fi-360 {\\b\\f1\\fs28\\cf1 Bold} plain {\\i italic}{\\footnote note text}\\par' +
        '\\pard\\trowd\\cellx2000\\cellx4000\\intbl a\\cell b\\cell\\row\\pard second\\tab tabbed\\line broken\\par' +
        '\\sect\\sectd\\pgwsxn10000\\pghsxn8000 next section\\par',
      '{\\stylesheet{\\s0 Normal;}{\\s1\\b heading 1;}}',
    ),
  );
  // The source has a footnote and no separators: the two that are made for it are not part of the comparison.
  const built = await build(source);
  assert.deepEqual(Array.from(built.result.footnotes?.items.keys() || []).slice(0, 2), [-1, 0]);
  built.result.footnotes?.items.delete(-1);
  built.result.footnotes?.items.delete(0);
  assert.equal(firstDifference(convertedPicture(source), convertedPicture(built.result)), undefined);
  const texts = built.result.body.blocks.filter((block): block is Paragraph => block.type === 'paragraph').map(paragraphText);
  assert.ok(texts.indexOf('second\ttabbed\nbroken') >= 0, JSON.stringify(texts));
  assert.equal(built.result.headers.size, source.headers.size);
  assert.ok(source.headers.size > 0 && source.footnotes !== undefined);
});

test('a document read from binary doc comes back the same', async () => {
  const source = simpleDoc('Merhaba dünya\rİkinci paragraf: ğüşıöç\r\rSon\r').read();
  const built = await same(source);
  assert.deepEqual(
    built.result.body.blocks.filter((block): block is Paragraph => block.type === 'paragraph').map(paragraphText),
    ['Merhaba dünya', 'İkinci paragraf: ğüşıöç', '', 'Son'],
  );
});

test('styles, numbering, fonts and settings are written from the model', async () => {
  const source = model([para('one', { style: 'ListPara', numbering: { numId: 2, level: 0 } }), para('two', { numbering: { numId: 2, level: 0 } }), para('sub', { numbering: { numId: 2, level: 1 } })]);
  source.styles.defaults = { run: { fonts: { ascii: 'Times New Roman', hAnsi: 'Times New Roman' }, size: 20, lang: { val: 'tr-TR' } }, para: { widowControl: true } };
  source.styles.styles.set('Normal', { id: 'Normal', type: 'paragraph', isDefault: true, name: 'Normal', quickFormat: true });
  source.styles.styles.set('ListPara', { id: 'ListPara', type: 'paragraph', isDefault: false, name: 'List Paragraph', basedOn: 'Normal', para: { indent: { left: 720 } }, run: { italic: true } });
  source.styles.styles.set('Emph', { id: 'Emph', type: 'character', isDefault: false, name: 'Emphasis', custom: true, run: { bold: true, color: { rgb: 'FF0000' } } });
  source.styles.defaultIds = { paragraph: 'Normal' };
  source.numbering.abstracts.set(1, {
    id: 1,
    multiLevelType: 'multilevel',
    levels: new Map([
      [0, { level: 0, start: 1, format: 'decimal', text: '%1.', align: 'left', para: { indent: { left: 360, hanging: 360 } } }],
      [1, { level: 1, start: 1, format: 'lowerLetter', text: '%1.%2)', align: 'left' }],
    ]),
  });
  source.numbering.nums.set(2, { id: 2, abstractId: 1, overrides: new Map() });
  source.fonts.set('Times New Roman', { name: 'Times New Roman', altName: 'Times', panose: '02020603050405020304', charset: 'A2', family: 'roman', pitch: 'variable' });
  source.fonts.set('Symbol', { name: 'Symbol', charset: '02' });
  source.settings.defaultTabStop = 708;
  source.settings.evenAndOddHeaders = true;
  source.settings.mirrorMargins = true;
  source.settings.autoHyphenation = true;
  source.settings.trackRevisions = true;
  source.settings.compat.set('useFELayout', true);
  source.settings.compatSettings.set('compatibilityMode', '11');
  source.settings.themeFontLang = { val: 'tr-TR' };
  const built = await same(source);

  const settings = await built.part('/word/settings.xml');
  assert.ok(
    settings.indexOf(
      '<w:mirrorMargins/><w:trackRevisions/><w:defaultTabStop w:val="708"/><w:autoHyphenation/><w:evenAndOddHeaders/>' +
        '<w:compat><w:useFELayout/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="11"/></w:compat><w:themeFontLang w:val="tr-TR"/></w:settings>',
    ) > 0,
    settings,
  );
  const styles = await built.part('/word/styles.xml');
  assert.ok(styles.indexOf('<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/><w:lang w:val="tr-TR"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:widowControl/></w:pPr></w:pPrDefault></w:docDefaults>') > 0, styles);
  const numbering = await built.part('/word/numbering.xml');
  assert.ok(numbering.indexOf('</w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>') > 0, numbering);
  const fonts = await built.part('/word/fontTable.xml');
  assert.ok(fonts.indexOf('<w:font w:name="Times New Roman"><w:altName w:val="Times"/><w:panose1 w:val="02020603050405020304"/><w:charset w:val="A2"/><w:family w:val="roman"/><w:pitch w:val="variable"/></w:font>') > 0, fonts);

  // The labels of the list are the same in both.
  const labels = (doc: DocModel) => {
    const counters = new NumberingCounters(doc.numbering, doc.styles);
    return doc.body.blocks.map((block) => (block.type === 'paragraph' && block.props.numbering ? counters.next(block.props.numbering.numId || 0, block.props.numbering.level || 0)?.text : undefined));
  };
  assert.deepEqual(labels(built.result), ['1.', '2.', '2.a)']);
  assert.deepEqual(labels(built.result), labels(source));
});

test('headers and footers: made-up ids become relationships, sections refer to them', async () => {
  const header = story('doc:header', [para('head')]);
  const footer = story('doc:footer', [para('foot')]);
  const first = story('doc:first', [para('first page')]);
  const source = model([para('one', { section: { type: 'nextPage', headers: { default: 'hdr:0:default', first: 'hdr:0:first' }, footers: { default: 'ftr:0:default' }, titlePage: true } }), para('two')], {
    // The second section shows the header of the first.
    headers: new Map([
      ['hdr:0:default', header],
      ['hdr:0:first', first],
      ['hdr:1:default', header],
    ]),
    footers: new Map([['ftr:0:default', footer]]),
  });
  source.body.section = { headers: { default: 'hdr:1:default' } };
  const built = await same(source);
  assert.deepEqual(built.pkg.partNames().filter((name) => /header|footer/.test(name)).sort(), ['/word/footer1.xml', '/word/header1.xml', '/word/header2.xml']);
  const rels = await built.pkg.relationships('/word/document.xml');
  assert.deepEqual(
    rels.filter((rel) => /\/(header|footer)$/.test(rel.type)).map((rel) => [rel.id, rel.part]),
    [
      ['rId1', '/word/header1.xml'],
      ['rId2', '/word/header2.xml'],
      ['rId3', '/word/header1.xml'],
      ['rId4', '/word/footer1.xml'],
    ],
  );
  // Ids that share a story share it after reading too.
  assert.equal(built.result.headers.get('rId1'), built.result.headers.get('rId3'));
  assert.equal(built.pkg.contentTypeOf('/word/header1.xml'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');
  const document = await built.part('/word/document.xml');
  assert.ok(document.indexOf('<w:sectPr><w:headerReference w:type="default" r:id="rId1"/><w:headerReference w:type="first" r:id="rId2"/><w:footerReference w:type="default" r:id="rId4"/><w:type w:val="nextPage"/><w:titlePg/></w:sectPr>') > 0, document);
});

test('notes get the separators Word expects; comments get an author', async () => {
  const source = model(
    [
      {
        type: 'paragraph',
        props: {},
        children: [
          { type: 'run', props: {}, content: [{ type: 'text', text: 'text' }, { type: 'footnoteReference', id: 1 }, { type: 'endnoteReference', id: 1 }, { type: 'commentReference', id: 0 }] },
        ],
      },
    ],
    {
      footnotes: collection('doc:footnotes', [[1, { id: 1, blocks: [para('a footnote')] }]]),
      endnotes: collection('doc:endnotes', [
        [-1, { id: -1, kind: 'separator', blocks: [para('custom separator')] }],
        [1, { id: 1, blocks: [para('an endnote')] }],
      ]),
      comments: collection('doc:comments', [[0, { id: 0, initials: 'AB', date: '2024-01-02T03:04:05Z', blocks: [para('a comment')] }]]),
    },
  );
  const built = await build(source);
  const footnotes = await built.part('/word/footnotes.xml');
  assert.ok(
    footnotes.indexOf(
      '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>' +
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
        '<w:footnote w:id="1"><w:p><w:r><w:t>a footnote</w:t></w:r></w:p></w:footnote></w:footnotes>',
    ) > 0,
    footnotes,
  );
  // The separator of the model is kept; only the missing one is made.
  const endnotes = built.result.endnotes;
  assert.ok(endnotes);
  assert.deepEqual(Array.from(endnotes.items.values()).map((note) => [note.id, note.kind]), [[0, 'continuationSeparator'], [-1, 'separator'], [1, undefined]]);
  const settings = await built.part('/word/settings.xml');
  assert.ok(settings.indexOf('<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr><w:endnotePr><w:endnote w:id="0"/><w:endnote w:id="-1"/></w:endnotePr>') > 0, settings);
  const comments = await built.part('/word/comments.xml');
  assert.ok(comments.indexOf('<w:comment w:id="0" w:author="" w:date="2024-01-02T03:04:05Z" w:initials="AB">') > 0, comments);
  assert.deepEqual(
    (await built.pkg.relationships('/word/document.xml')).map((rel) => rel.type.slice(rel.type.lastIndexOf('/') + 1)).sort(),
    ['comments', 'endnotes', 'footnotes', 'settings', 'styles'],
  );
});

test('hyperlinks: only a target the engine would follow becomes a relationship', async () => {
  const link = (wrapper: Wrapper, text: string): Paragraph => ({ type: 'paragraph', props: {}, children: [{ type: 'run', props: {}, content: [{ type: 'text', text }], wrap: wrapper }] });
  const source = model([
    link({ kind: 'hyperlink', target: 'https://example.com/a?b=1&c=2', tooltip: 'tip' }, 'web'),
    link({ kind: 'hyperlink', target: 'https://example.com/a?b=1&c=2' }, 'web again'),
    link({ kind: 'hyperlink', target: 'javascript:alert(1)' }, 'script'),
    link({ kind: 'hyperlink', target: 'file:///etc/passwd' }, 'file'),
    link({ kind: 'hyperlink', target: '#Bookmark1' }, 'internal'),
    link({ kind: 'hyperlink', anchor: 'Bookmark2' }, 'anchor'),
    link({ kind: 'hyperlink', target: ' mailto:someone@example.com ' }, 'mail'),
  ]);
  const built = await build(source);
  const rels = (await built.pkg.relationships('/word/document.xml')).filter((rel) => rel.type.endsWith('/hyperlink'));
  assert.deepEqual(rels.map((rel) => [rel.target, rel.external]), [
    ['https://example.com/a?b=1&c=2', true],
    ['mailto:someone@example.com', true],
  ]);
  const document = await built.part('/word/document.xml');
  const body = document.slice(document.indexOf('<w:body>'));
  assert.equal(
    body,
    `<w:body><w:p><w:hyperlink r:id="${rels[0].id}" w:tooltip="tip"><w:r><w:t>web</w:t></w:r></w:hyperlink></w:p><w:p><w:hyperlink r:id="${rels[0].id}"><w:r><w:t>web again</w:t></w:r></w:hyperlink></w:p>` +
      '<w:p><w:r><w:t>script</w:t></w:r></w:p><w:p><w:r><w:t>file</w:t></w:r></w:p>' +
      '<w:p><w:hyperlink w:anchor="Bookmark1"><w:r><w:t>internal</w:t></w:r></w:hyperlink></w:p><w:p><w:hyperlink w:anchor="Bookmark2"><w:r><w:t>anchor</w:t></w:r></w:hyperlink></w:p>' +
      `<w:p><w:hyperlink r:id="${rels[1].id}"><w:r><w:t>mail</w:t></w:r></w:hyperlink></w:p></w:body></w:document>`,
  );
  assert.ok(document.indexOf('javascript') < 0 && document.indexOf('passwd') < 0);
});

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 2, 0, 0, 0, 1, 8, 2, 0, 0, 0]);
const EMF = new Uint8Array([1, 0, 0, 0, 88, 0, 0, 0, 1, 2, 3]);

const standIn = (insets?: string[]): DrawingItem['xml'] => ({
  type: 'element',
  name: 'doc:drawing',
  attrs: new Map(),
  children: insets ? [{ type: 'element', name: 'wps:bodyPr', attrs: new Map([['lIns', insets[0]], ['tIns', insets[1]], ['rIns', insets[2]], ['bIns', insets[3]]]), children: [] }] : [],
});

const inRun = (...items: DrawingItem[]): Paragraph => ({ type: 'paragraph', props: {}, children: [{ type: 'run', props: {}, content: items }] });

test('drawings without a tree are written as standard pictures, text boxes and groups', async () => {
  const png = { rId: 'doc:image1', target: '/doc/media/image1.png' };
  const emf = { rId: 'doc:image2', target: '/doc/media/image2.emf' };
  const inline: DrawingItem = { type: 'drawing', placement: 'inline', kind: 'picture', width: 914400, height: 457200, description: 'alt & "text"', image: png, images: [{ image: png, crop: { left: 1000, bottom: -2000 } }], crop: { left: 1000, bottom: -2000 }, xml: standIn() };
  const again: DrawingItem = { type: 'drawing', placement: 'inline', kind: 'picture', width: 100, height: 200, image: png, images: [{ image: png }], xml: standIn() };
  const box: DrawingItem = {
    type: 'drawing',
    placement: 'anchor',
    kind: 'shape',
    width: 1828800,
    height: 914400,
    name: 'Text Box 1',
    anchor: { horizontal: { relativeFrom: 'page', offset: 360000 }, vertical: { relativeFrom: 'paragraph', align: 'top' }, wrap: 'square', wrapText: 'left', behindDoc: true, distL: 114300, distR: 114300 },
    transform: { rotation: 5400000, flipH: true },
    textBoxes: [[para('in the box'), { type: 'table', props: {}, grid: [500], rows: [{ props: {}, cells: [{ props: {}, blocks: [para('cell')] }] }] }]],
    xml: standIn(['0', '45720', '91440', '0']),
  };
  const group: DrawingItem = {
    type: 'drawing',
    placement: 'anchor',
    kind: 'group',
    width: 2000000,
    height: 1000000,
    anchor: { horizontal: { relativeFrom: 'column', offset: 0 }, vertical: { relativeFrom: 'paragraph', offset: -100 }, wrap: 'tight' },
    image: emf,
    images: [{ image: emf }, { image: png }],
    textBoxes: [[para('first box')], [para('second box')]],
    xml: standIn(),
  };
  const plain: DrawingItem = { type: 'drawing', placement: 'anchor', kind: 'shape', width: 300, height: 400, anchor: { wrap: 'topAndBottom' }, xml: standIn() };
  const noImage: DrawingItem = { type: 'drawing', placement: 'inline', kind: 'picture', width: 5, height: 6, xml: standIn() };
  const source = model([inRun(inline, again), inRun(box), inRun(group), inRun(plain, noImage)], {
    media: new Map([
      ['/doc/media/image1.png', { bytes: PNG, type: 'image/png' }],
      ['/doc/media/image2.emf', { bytes: EMF, type: 'image/x-emf' }],
    ]),
  });
  const built = await same(source);

  // The pictures are parts of the package, as they were, under the same names.
  assert.deepEqual(await built.pkg.readPart('/word/media/image1.png'), PNG);
  assert.deepEqual(await built.pkg.readPart('/word/media/image2.emf'), EMF);
  assert.equal(built.pkg.contentTypeOf('/word/media/image1.png'), 'image/png');
  assert.equal(built.pkg.contentTypeOf('/word/media/image2.emf'), 'image/x-emf');
  // One relationship for a picture, however often it is used.
  const rels = (await built.pkg.relationships('/word/document.xml')).filter((rel) => rel.type.endsWith('/image'));
  assert.deepEqual(rels.map((rel) => rel.target), ['media/image1.png', 'media/image2.emf']);

  const document = await built.part('/word/document.xml');
  assert.ok(document.indexOf('descr="alt &amp; &quot;text&quot;"') > 0);
  assert.ok(document.indexOf('<a:srcRect l="1000" b="-2000"/>') > 0);
  assert.ok(document.indexOf('<wp:positionH relativeFrom="page"><wp:posOffset>360000</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:align>top</wp:align></wp:positionV>') > 0);
  assert.ok(document.indexOf('<wp:wrapSquare wrapText="left"/>') > 0 && document.indexOf('behindDoc="1"') > 0);
  assert.ok(document.indexOf('<a:xfrm rot="5400000" flipH="1">') > 0);
  assert.ok(document.indexOf('<wps:bodyPr lIns="0" tIns="45720" rIns="91440" bIns="0"/>') > 0);
  assert.ok(document.indexOf('<wp:wrapTight wrapText="bothSides"><wp:wrapPolygon edited="0"><wp:start x="0" y="0"/>') > 0);
  // Drawing ids are unique in the document.
  const ids = document.match(/<wp:docPr id="\d+"/g) || [];
  assert.equal(new Set(ids).size, 6);

  const read = (block: Block, index = 0): DrawingItem => {
    assert.ok(block.type === 'paragraph' && block.children[0].type === 'run');
    const item = block.children[0].content[index];
    assert.ok(item.type === 'drawing');
    return item;
  };
  const blocks = built.result.body.blocks;
  assert.deepEqual([read(blocks[0]).kind, read(blocks[1]).kind, read(blocks[2]).kind, read(blocks[3]).kind, read(blocks[3], 1).kind], ['picture', 'shape', 'group', 'shape', 'picture']);
  assert.equal(read(blocks[0]).image?.target, '/word/media/image1.png');
  assert.equal(read(blocks[1]).textBoxes?.[0].length, 2);
  assert.deepEqual(read(blocks[2]).images?.map((item) => item.image.target), ['/word/media/image2.emf', '/word/media/image1.png']);
  assert.equal(read(blocks[3], 1).image, undefined);
});

test('a text box with a picture keeps both; a picture that is not available leaves a drawing without an image', async () => {
  const png = { rId: 'x', target: '/doc/media/image1.png' };
  const gone = { rId: 'y', target: '/doc/media/gone.png' };
  const both: DrawingItem = { type: 'drawing', placement: 'anchor', kind: 'shape', width: 10, height: 20, anchor: {}, image: png, images: [{ image: png }], textBoxes: [[para('text')]], xml: standIn() };
  const missing: DrawingItem = { type: 'drawing', placement: 'inline', kind: 'picture', width: 10, height: 20, image: gone, images: [{ image: gone }], xml: standIn() };
  const built = await build(model([inRun(both, missing)], { media: new Map([['/doc/media/image1.png', { bytes: PNG, type: 'image/png' }]]) }));
  const block = built.result.body.blocks[0];
  assert.ok(block.type === 'paragraph' && block.children[0].type === 'run');
  const [first, second] = block.children[0].content;
  assert.ok(first.type === 'drawing' && second.type === 'drawing');
  assert.deepEqual([first.kind, first.image?.target, first.textBoxes?.length], ['shape', '/word/media/image1.png', 1]);
  assert.deepEqual([second.kind, second.image, second.width], ['picture', undefined, 10]);
});

test('tracked changes, fields, bookmarks and tables of a model without a package', async () => {
  const inserted: Wrapper = { kind: 'ins', id: 1, author: 'A', date: '2024-01-01T00:00:00Z' };
  // (With an id: a revision without one is given one, see the tests of the third review.)
  const deleted: Wrapper = { kind: 'del', id: 2, author: 'B' };
  const source = model([
    {
      type: 'paragraph',
      props: { align: 'both', spacing: { before: 120, line: 276, lineRule: 'auto' }, tabs: [{ pos: 4536, align: 'center' }], markProps: { bold: true } },
      children: [
        { type: 'bookmarkStart', id: '0', name: '_Toc1' },
        { type: 'run', props: { bold: true, fonts: { ascii: 'Arial', hAnsi: 'Arial', cs: 'Arial' }, size: 24, underline: { style: 'single' }, color: { rgb: '0000FF' } }, content: [{ type: 'text', text: ' kept space ' }, { type: 'tab' }, { type: 'break', kind: 'line' }, { type: 'symbol', font: 'Symbol', char: '' }] },
        { type: 'bookmarkEnd', id: '0' },
        { type: 'run', props: {}, content: [{ type: 'text', text: 'new' }], wrap: inserted },
        { type: 'run', props: {}, content: [{ type: 'deletedText', text: 'old' }], wrap: deleted },
        { type: 'run', props: {}, content: [{ type: 'fieldChar', charType: 'begin', field: { instruction: ' PAGE ', result: '1' } }] },
        { type: 'run', props: {}, content: [{ type: 'instrText', text: ' PAGE ' }], fieldCode: true },
        { type: 'run', props: {}, content: [{ type: 'fieldChar', charType: 'separate' }], fieldCode: true },
        { type: 'run', props: {}, content: [{ type: 'text', text: '1' }] },
        { type: 'run', props: {}, content: [{ type: 'fieldChar', charType: 'end' }] },
        { type: 'run', props: {}, content: [{ type: 'break', kind: 'page' }, { type: 'pageNumber' }, { type: 'noBreakHyphen' }] },
      ],
    },
    {
      type: 'table',
      props: { width: { type: 'dxa', value: 5000 }, borders: { top: { style: 'single', size: 4, color: { auto: true } }, insideV: { style: 'single', size: 4 } }, look: { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true }, layout: 'fixed' },
      grid: [2500, 2500],
      rows: [
        { props: { header: true, cantSplit: true, height: 300, heightRule: 'atLeast' }, cells: [{ props: { width: { type: 'dxa', value: 5000 }, gridSpan: 2, shading: { pattern: 'clear', fill: { rgb: 'D9D9D9' } } }, blocks: [para('merged')] }] },
        {
          props: {},
          cells: [
            { props: { width: { type: 'dxa', value: 2500 }, vMerge: 'restart', vAlign: 'center' }, blocks: [para('a')] },
            { props: { width: { type: 'dxa', value: 2500 }, borders: { left: { style: 'double', size: 6 } }, margins: { left: { type: 'dxa', value: 50 } } }, blocks: [para('b'), { type: 'table', props: {}, grid: [100], rows: [{ props: {}, cells: [{ props: {}, blocks: [para('nested')] }] }] }, para('')] },
          ],
        },
        { props: {}, cells: [{ props: { vMerge: 'continue' }, blocks: [para('')] }, { props: {}, blocks: [para('c')] }] },
      ],
    },
    para('after'),
  ]);
  source.body.section = { page: { width: 11906, height: 16838, orientation: 'portrait' }, columns: { count: 2, space: 708 }, pageNumbering: { format: 'lowerRoman', start: 3 } };
  const built = await same(source);
  const document = await built.part('/word/document.xml');
  assert.ok(document.indexOf('<w:t xml:space="preserve"> kept space </w:t><w:tab/><w:br/><w:sym w:font="Symbol" w:char="F0B7"/>') > 0, document);
  assert.ok(document.indexOf('<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>') > 0, document);
});

test('a model read from docx is built again: its pictures come from the given source, theme values become plain ones', async () => {
  const bytes = await buildSpec({
    body:
      '<w:p><w:pPr><w:pStyle w:val="H"/></w:pPr><w:r><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/><w:color w:val="000000" w:themeColor="accent1"/></w:rPr><w:t>themed</w:t></w:r>' +
      '<w:hyperlink r:id="rIdL"><w:r><w:t>link</w:t></w:r></w:hyperlink><w:hyperlink r:id="rIdBad"><w:r><w:t>bad</w:t></w:r></w:hyperlink>' +
      '<w:r><w:drawing><wp:inline><wp:extent cx="100" cy="200"/><wp:docPr id="1" name="P"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdI"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>' +
      '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>',
    styles: '<w:style w:type="paragraph" w:styleId="H"><w:name w:val="heading 1"/><w:rPr><w:color w:themeColor="accent2"/></w:rPr></w:style>',
    theme: THEME,
    headers: { rIdH: '<w:p><w:r><w:t>head</w:t></w:r></w:p>' },
    rels: [
      ['rIdL', 'hyperlink', 'https://example.com/', true],
      ['rIdBad', 'hyperlink', 'javascript:alert(1)', true],
      ['rIdI', 'image', 'media/image7.png'],
    ],
    parts: [['word/media/image7.png', PNG]],
  });
  const pkg = await OpcPackage.open(bytes);
  const source = await readDocx(pkg);
  const report = emptyWriteReport();
  const built = await build(source, { pictures: pkg, report });
  assert.deepEqual(await built.pkg.readPart('/word/media/image7.png'), PNG);
  assert.equal(built.pkg.hasPart('/word/theme/theme1.xml'), false);
  const document = await built.part('/word/document.xml');
  // The fonts and colours the theme gave, stated plainly.
  assert.ok(document.indexOf('<w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:color w:val="4472C4"/>') > 0, document);
  assert.ok((await built.part('/word/styles.xml')).indexOf('<w:color w:val="ED7D31"/>') > 0);
  // The kept drawing still finds its picture, the safe link its target; the other link is text.
  assert.ok(document.indexOf('<w:hyperlink r:id="rIdL">') > 0 && document.indexOf('rIdBad') < 0 && document.indexOf('<w:r><w:t>bad</w:t></w:r>') > 0);
  const block = built.result.body.blocks[0];
  assert.ok(block.type === 'paragraph');
  const last = block.children[block.children.length - 1];
  assert.ok(last.type === 'run' && last.content[0].type === 'drawing');
  assert.equal(last.content[0].image?.target, '/word/media/image7.png');
  assert.equal(built.result.headers.size, 1);
  assert.equal(JSON.stringify((await built.pkg.relationships('/word/document.xml')).map((rel) => rel.target)).indexOf('javascript'), -1);
  assert.ok(report.parts.indexOf('/word/media/image7.png') >= 0);
  // What was not carried over is in the report: the link that became text, and the theme.
  assert.deepEqual(report.warnings, [
    '1 link(s) with an address that is not followed were written as plain text',
    'The theme was not carried over: theme colours and theme fonts were written as the plain values they stand for',
  ]);
  assert.equal(report.droppedCharacters, 0);
});

test('the report counts what a new package does not carry over; a model that loses nothing gives no warning', async () => {
  const link = (target: string, text: string): Paragraph => ({ type: 'paragraph', props: {}, children: [{ type: 'run', props: {}, content: [{ type: 'text', text }], wrap: { kind: 'hyperlink', target } }] });
  const png = { rId: 'x', target: '/doc/media/image1.png' };
  const gone = { rId: 'y', target: '/doc/media/gone.png' };
  const media = (): DocModel['media'] => new Map([['/doc/media/image1.png', { bytes: PNG, type: 'image/png' }]]);
  const picture: DrawingItem = { type: 'drawing', placement: 'inline', kind: 'picture', width: 10, height: 20, image: png, images: [{ image: png }], xml: standIn() };

  const clean = emptyWriteReport();
  await build(model([para('Plain text'), inRun(picture), link('https://example.com/', 'web'), link('#Mark', 'inside')], { media: media() }), { report: clean });
  assert.deepEqual(clean.warnings, []);
  assert.equal(clean.droppedCharacters, 0);

  const bare: DrawingItem = { type: 'drawing', placement: 'anchor', kind: 'shape', width: 10, height: 20, anchor: {}, xml: standIn() };
  const box: DrawingItem = { type: 'drawing', placement: 'anchor', kind: 'shape', width: 10, height: 20, anchor: {}, textBoxes: [[para('in the box'), inRun(bare)]], xml: standIn() };
  const group: DrawingItem = { type: 'drawing', placement: 'anchor', kind: 'group', width: 10, height: 20, anchor: {}, images: [{ image: png }], textBoxes: [[para('one')], [para('two')]], xml: standIn() };
  const missing: DrawingItem = { type: 'drawing', placement: 'inline', kind: 'picture', width: 10, height: 20, image: gone, images: [{ image: gone }], xml: standIn() };
  const untargeted: DrawingItem = { type: 'drawing', placement: 'inline', kind: 'picture', width: 10, height: 20, image: { rId: 'z' }, xml: standIn() };
  const report = emptyWriteReport();
  await build(
    model([para('A\u0001B\u0002C'), inRun(bare, box, group, missing, untargeted, picture), link('javascript:alert(1)', 'script'), link('file:///etc/passwd', 'file'), link('https://example.com/', 'web')], {
      media: media(),
    }),
    { report },
  );
  assert.deepEqual(report.warnings, [
    // The bare shape twice (once inside the text box) and the text box itself.
    '3 shape(s) were written as plain rectangles: their outline, fill and geometry were not carried over',
    '1 group(s) of shapes were written with every part over the whole area of the group, without outlines and fills',
    '2 picture(s) whose data is not available were written as empty frames',
    '2 link(s) with an address that is not followed were written as plain text',
  ]);
  assert.equal(report.droppedCharacters, 2);
});

test('a strict model is built with the strict namespaces', async () => {
  const source = model([para('x', { cnfStyle: '100000000000' })], { strict: true });
  const built = await build(source);
  assert.equal(built.result.strict, true);
  const document = await built.part('/word/document.xml');
  assert.ok(document.indexOf('xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"') > 0);
  assert.ok(document.indexOf('<w:cnfStyle w:firstRow="1"/>') > 0);
  const rels = await built.pkg.relationships('/word/document.xml');
  assert.ok(rels.every((rel) => rel.type.indexOf('http://purl.oclc.org/ooxml/officeDocument/relationships/') === 0));
});
