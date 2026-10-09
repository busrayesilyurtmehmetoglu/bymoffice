// A copy takes along what the font table of its document says of the fonts it
// names; a document that is pasted into and does not know them gets the
// entries, in its model and in the file written from it.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import type { DocModel } from '../src/model/doc/types';
import { OpcPackage } from '../src/opc';
import { DocHistory, applyDocOp, fragmentOf, insertFragment, type DocFragment, type DocPosition, type DocRange } from '../src/ops/doc/index';
import { buildDocx, type DocxSpec } from './docx-fixtures';

const STYLES =
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:rPr><w:rFonts w:ascii="Styled Face" w:hAnsi="Styled Face"/></w:rPr></w:style>';
const NUMBERING =
  '<w:abstractNum w:abstractNumId="4"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/><w:rPr><w:rFonts w:ascii="Bullet Face" w:hAnsi="Bullet Face"/></w:rPr></w:lvl></w:abstractNum>' +
  '<w:num w:numId="7"><w:abstractNumId w:val="4"/></w:num>';
const SECT = '<w:sectPr><w:pgSz w:w="9000" w:h="9000"/></w:sectPr>';

const font = (name: string, inner = '<w:charset w:val="A2"/><w:family w:val="roman"/><w:pitch w:val="variable"/>') => `<w:font w:name="${name}">${inner}</w:font>`;
const run = (text: string, face?: string) => `<w:r>${face ? `<w:rPr><w:rFonts w:ascii="${face}" w:hAnsi="${face}"/></w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const para = (content: string, pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${content}</w:p>`;

interface Opened {
  pkg: OpcPackage;
  doc: DocModel;
}

async function open(spec: DocxSpec): Promise<Opened> {
  const full: DocxSpec = { ...spec };
  if (full.styles === undefined) full.styles = STYLES;
  full.body = (full.body || '') + SECT;
  const pkg = await OpcPackage.open(await buildDocx(full));
  return { pkg, doc: await readDocx(pkg) };
}

const at = (path: number, offset: number): DocPosition => ({ story: 'body', path: [path], offset });
const range = (a: DocPosition, b: DocPosition): DocRange => ({ anchor: a, focus: b });
const decoder = new TextDecoder();

function paste(doc: DocModel, fragment: DocFragment, where: DocPosition): DocHistory {
  const history = new DocHistory(doc);
  history.applyEdit(insertFragment(doc, where, fragment), undefined);
  history.seal();
  return history;
}

async function written(opened: Opened, part = '/word/fontTable.xml'): Promise<string | undefined> {
  await writeDocx(opened.pkg, opened.doc);
  const bytes = await opened.pkg.readPart(part);
  return bytes ? decoder.decode(bytes) : undefined;
}

async function reread(opened: Opened): Promise<DocModel> {
  await writeDocx(opened.pkg, opened.doc);
  return readDocx(await OpcPackage.open(await opened.pkg.save()));
}

const SOURCE: DocxSpec = {
  body: para(run('plain ') + run('fancy', 'Fancy Face') + run(' & more', 'Rock & Roll')) + para(run('quoted'), '<w:pStyle w:val="Quote"/>') + para(run('item'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>') + para(run('end')),
  numbering: NUMBERING,
  fontTable:
    font('Fancy Face', '<w:altName w:val="Fancy Alt"/><w:panose1 w:val="02020603050405020304"/><w:charset w:val="A2"/><w:family w:val="roman"/><w:pitch w:val="variable"/><w:embedRegular r:id="rId9"/>') +
    font('Rock &amp; Roll') +
    font('Styled Face') +
    font('Bullet Face') +
    font('Unused Face'),
};

test('a copy holds what its document says of the fonts it names, and of no other', async () => {
  const source = await open(SOURCE);
  const fragment = fragmentOf(source.doc, range(at(0, 0), at(0, 11)));
  assert.ok(fragment.fonts instanceof Map);
  assert.deepEqual([...fragment.fonts.keys()], ['Fancy Face']);
  const info = fragment.fonts.get('Fancy Face')!;
  assert.equal(info.altName, 'Fancy Alt');
  assert.equal(info.panose, '02020603050405020304');
  assert.equal(info.family, 'roman');
  // The embedded file is of the package it is in.
  assert.equal(info.embedded, undefined);
  assert.ok(source.doc.fonts.get('Fancy Face')!.embedded, 'the document keeps what it had');
});

test('a paste into a document that does not know the font adds its entry; the file names it and reads back the same', async () => {
  const source = await open(SOURCE);
  const target = await open({ body: para(run('here')), fontTable: font('Own Face') });
  const table = target.doc.fonts;
  const fragment = fragmentOf(source.doc, range(at(0, 0), at(0, 18)));
  paste(target.doc, fragment, at(0, 4));
  // The table stays the object views hold.
  assert.equal(target.doc.fonts, table);
  assert.deepEqual([...target.doc.fonts.keys()], ['Own Face', 'Fancy Face', 'Rock & Roll']);

  const text = (await written(target))!;
  assert.equal(text.split('<w:font ').length - 1, 3);
  assert.ok(text.indexOf('<w:font w:name="Own Face">') < text.indexOf('<w:font w:name="Fancy Face">'));
  assert.ok(text.includes('<w:font w:name="Rock &amp; Roll">'));
  assert.ok(!text.includes('embedRegular'), 'no reference to a file the package does not have');
  const again = await reread(target);
  assert.deepEqual([...again.fonts.keys()], ['Own Face', 'Fancy Face', 'Rock & Roll']);
  assert.deepEqual(again.fonts.get('Fancy Face'), { name: 'Fancy Face', altName: 'Fancy Alt', panose: '02020603050405020304', charset: 'A2', family: 'roman', pitch: 'variable' });
});

test('undo takes the entry out again and the file is what it was; redo brings it back', async () => {
  const source = await open(SOURCE);
  const target = await open({ body: para(run('here')), fontTable: font('Own Face') });
  const before = await target.pkg.readPart('/word/fontTable.xml');
  const history = paste(target.doc, fragmentOf(source.doc, range(at(0, 0), at(0, 11))), at(0, 4));
  assert.ok((await written(target))!.includes('Fancy Face'));
  history.undo();
  assert.deepEqual([...target.doc.fonts.keys()], ['Own Face']);
  await writeDocx(target.pkg, target.doc);
  assert.deepEqual(await target.pkg.readPart('/word/fontTable.xml'), before);
  history.redo();
  assert.deepEqual([...target.doc.fonts.keys()], ['Own Face', 'Fancy Face']);
  assert.ok((await written(target))!.includes('Fancy Face'));
});

test('a font the document knows keeps the entry of the document, and a paste into the document of the copy adds nothing', async () => {
  const source = await open(SOURCE);
  const target = await open({ body: para(run('here')), fontTable: font('Fancy Face', '<w:family w:val="swiss"/>') });
  const before = await target.pkg.readPart('/word/fontTable.xml');
  paste(target.doc, fragmentOf(source.doc, range(at(0, 0), at(0, 11))), at(0, 4));
  assert.deepEqual(target.doc.fonts.get('Fancy Face'), { name: 'Fancy Face', family: 'swiss' });
  await writeDocx(target.pkg, target.doc);
  assert.deepEqual(await target.pkg.readPart('/word/fontTable.xml'), before);

  const size = source.doc.fonts.size;
  paste(source.doc, fragmentOf(source.doc, range(at(0, 0), at(0, 11))), at(1, 0));
  assert.equal(source.doc.fonts.size, size);
});

test('the fonts of a style and of a list that come with the paste come too; those of a style the document has do not', async () => {
  const source = await open(SOURCE);
  const target = await open({ body: para(run('here')), styles: STYLES.replace('Styled Face', 'Other Face').replace('Styled Face', 'Other Face') });
  assert.equal(target.doc.fonts.size, 0);
  // The quoted paragraph and the list item, each with its end: the item stays an item where it is pasted.
  const fragment = fragmentOf(source.doc, range(at(1, 0), at(3, 0)));
  assert.deepEqual([...fragment.fonts!.keys()].sort(), ['Bullet Face', 'Styled Face']);
  paste(target.doc, fragment, at(0, 4));
  // "Quote" is a style of the document already, with a font of its own.
  assert.deepEqual([...target.doc.fonts.keys()], ['Bullet Face']);

  const fresh = await open({ body: para(run('here')), styles: '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' });
  paste(fresh.doc, fragment, at(0, 4));
  assert.deepEqual([...fresh.doc.fonts.keys()].sort(), ['Bullet Face', 'Styled Face']);
});

test('a document without a font table gets one, related to the document, when a paste brings a font', async () => {
  const source = await open(SOURCE);
  const target = await open({ body: para(run('here')) });
  assert.equal(target.pkg.hasPart('/word/fontTable.xml'), false);
  paste(target.doc, fragmentOf(source.doc, range(at(0, 0), at(0, 11))), at(0, 4));
  const text = (await written(target))!;
  assert.ok(text.includes('<w:font w:name="Fancy Face">'));
  const rels = decoder.decode((await target.pkg.readPart('/word/_rels/document.xml.rels'))!);
  assert.ok(/Type="[^"]*\/fontTable" Target="fontTable.xml"/.test(rels));
  const again = await reread(target);
  assert.deepEqual([...again.fonts.keys()], ['Fancy Face']);
});

test('an edit that brings no font leaves the font table of the file byte for byte, odd names in it included', async () => {
  const target = await open({
    body: para(run('here', 'Rock & Roll')),
    fontTable: font('Rock &amp; Roll') + font('Rock &amp; Roll', '<w:family w:val="swiss"/>') + font('A &#x41;&quot;B'),
  });
  const before = await target.pkg.readPart('/word/fontTable.xml');
  const history = new DocHistory(target.doc);
  history.applyEdit(insertFragment(target.doc, at(0, 4), fragmentOf(target.doc, range(at(0, 0), at(0, 2)))), undefined);
  history.seal();
  await writeDocx(target.pkg, target.doc);
  assert.deepEqual(await target.pkg.readPart('/word/fontTable.xml'), before);
});

test('a fragment made without fonts, or with entries that are none, is pasted as before', async () => {
  const source = await open(SOURCE);
  const target = await open({ body: para(run('here')) });
  const fragment = fragmentOf(source.doc, range(at(0, 0), at(0, 11)));
  delete fragment.fonts;
  paste(target.doc, fragment, at(0, 4));
  assert.equal(target.doc.fonts.size, 0);

  const odd = fragmentOf(source.doc, range(at(0, 0), at(0, 11)));
  odd.fonts = new Map([['Fancy Face', { name: 'Another' }]]);
  paste(target.doc, odd, at(0, 0));
  assert.equal(target.doc.fonts.size, 0);
});

test('setFonts: entries are put and taken by name, the inverse gives back what was, and what is no entry is refused', async () => {
  const { doc } = await open({ body: para(run('here')), fontTable: font('Own Face') });
  const own = doc.fonts.get('Own Face');
  const inverse = applyDocOp(doc, {
    type: 'setFonts',
    fonts: [
      ['New Face', { name: 'New Face', family: 'swiss' }],
      ['Own Face', null],
      ['New Face', { name: 'New Face', family: 'roman' }],
    ],
  });
  assert.deepEqual([...doc.fonts.entries()], [['New Face', { name: 'New Face', family: 'roman' }]]);
  applyDocOp(doc, inverse);
  assert.deepEqual([...doc.fonts.keys()], ['Own Face']);
  assert.equal(doc.fonts.get('Own Face'), own);

  assert.throws(() => applyDocOp(doc, { type: 'setFonts', fonts: [['', { name: '' }]] }), RangeError);
  assert.throws(() => applyDocOp(doc, { type: 'setFonts', fonts: [['A', { name: 'B' }]] }), RangeError);
  assert.throws(() => applyDocOp(doc, { type: 'setFonts', fonts: [['A', { name: 'A' }], ['B', 3 as never]] }), RangeError);
  assert.deepEqual([...doc.fonts.keys()], ['Own Face'], 'a refused operation changes nothing');
});

// ------------------------------------------------------------------ review

test('review: the font of a paragraph mark comes with the paragraph', async () => {
  const mark = '<w:rPr><w:rFonts w:ascii="Mark Face" w:hAnsi="Mark Face"/></w:rPr>';
  const source = await open({ body: para(run('one'), mark) + para(run('two')), fontTable: font('Mark Face') });
  const target = await open({ body: para(run('here')) });
  const fragment = fragmentOf(source.doc, range(at(0, 0), at(1, 0)));
  assert.deepEqual([...fragment.fonts!.keys()], ['Mark Face']);
  paste(target.doc, fragment, at(0, 4));
  assert.deepEqual([...target.doc.fonts.keys()], ['Mark Face']);
  assert.ok((await written(target))!.includes('<w:font w:name="Mark Face">'));
});

test('review: the measurer names the stand-ins of a font again once the font table has an entry more or less', async () => {
  const { CanvasTextMeasurer } = await import('../src/render/doc-measure');
  const fonts = new Map<string, { name: string; altName?: string; family?: string }>();
  const measurer = new CanvasTextMeasurer({ fonts, hasFamily: () => false });
  const before = measurer.cssFamilies('Fancy Face');
  assert.ok(!before.includes('Fancy Alt'));
  fonts.set('Fancy Face', { name: 'Fancy Face', altName: 'Fancy Alt', family: 'roman' });
  const after = measurer.cssFamilies('Fancy Face');
  assert.ok(after.includes('"Fancy Alt"'), after);
  assert.ok(/serif$/.test(after) && !/sans-serif$/.test(after), after);
  fonts.delete('Fancy Face');
  assert.equal(measurer.cssFamilies('Fancy Face'), before);
});

test('review: what a fragment says of a font is taken field by field: texts of a length, in an object of its own', async () => {
  const source = await open(SOURCE);
  const make = (info: unknown): DocFragment => {
    const fragment = fragmentOf(source.doc, range(at(0, 0), at(0, 11)));
    fragment.fonts = new Map([['Fancy Face', info as never]]);
    return fragment;
  };
  const shared = { name: 'Fancy Face', altName: { toString: () => 'x' }, family: null, pitch: 5, charset: 'A2', panose: 'x'.repeat(5000), embedded: new Map([['regular', 'rId1']]), extra: 1 };
  const target = await open({ body: para(run('here')) });
  paste(target.doc, make(shared), at(0, 4));
  const got = target.doc.fonts.get('Fancy Face');
  assert.deepEqual(got, { name: 'Fancy Face', charset: 'A2' });
  assert.notEqual(got, shared);
  const text = (await written(target))!;
  assert.ok(text.includes('<w:font w:name="Fancy Face"><w:charset w:val="A2"/></w:font>'), text);

  const long = await open({ body: para(run('here')) });
  const fragment = fragmentOf(source.doc, range(at(0, 0), at(0, 11)));
  for (const block of [fragment.fonts!]) block.set('Fancy Face', { name: 'Fancy Face'.repeat(100) });
  paste(long.doc, fragment, at(0, 4));
  assert.equal(long.doc.fonts.size, 0);

  assert.throws(() => applyDocOp(target.doc, { type: 'setFonts', fonts: [['A', { name: 'A', altName: 5 as never }]] }), RangeError);
  assert.throws(() => applyDocOp(target.doc, { type: 'setFonts', fonts: [['A', { name: 'A', embedded: {} as never }]] }), RangeError);
});

test('review: a font the file names under another prefix of the same namespace is a font the file has', async () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const table = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts xmlns:w="${W}" xmlns:x="${W}"><x:font x:name="Own Face"/></w:fonts>`;
  const source = await open(SOURCE);
  const target = await open({ body: para(run('here', 'Own Face')), parts: [['word/fontTable.xml', table]], fontTable: '' });
  assert.deepEqual([...target.doc.fonts.keys()], ['Own Face']);
  const before = await target.pkg.readPart('/word/fontTable.xml');
  const history = new DocHistory(target.doc);
  history.applyEdit(insertFragment(target.doc, at(0, 4), fragmentOf(target.doc, range(at(0, 0), at(0, 2)))), undefined);
  history.seal();
  await writeDocx(target.pkg, target.doc);
  assert.deepEqual(await target.pkg.readPart('/word/fontTable.xml'), before);
  paste(target.doc, fragmentOf(source.doc, range(at(0, 0), at(0, 11))), at(0, 0));
  const text = (await written(target))!;
  assert.equal(text, table.replace('</w:fonts>', '') + text.slice(table.length - '</w:fonts>'.length));
  assert.equal(text.split('Own Face').length - 1, 1);
  assert.ok(text.includes('<w:font w:name="Fancy Face">'));
});

test('review: fonts added to a part with a byte order mark and CRLF line ends change no other byte of it', async () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const head = `﻿<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:fonts xmlns:w="${W}">\r\n<w:font w:name="Own Face"/>\r\n`;
  const tail = '</w:fonts>\r\n';
  const encode = (text: string) => new TextEncoder().encode(text);
  const source = await open(SOURCE);
  const target = await open({ body: para(run('here')), parts: [['word/fontTable.xml', encode(head + tail)]], fontTable: '' });
  paste(target.doc, fragmentOf(source.doc, range(at(0, 0), at(0, 11))), at(0, 4));
  await writeDocx(target.pkg, target.doc);
  const bytes = (await target.pkg.readPart('/word/fontTable.xml'))!;
  const text = new TextDecoder('utf-8', { ignoreBOM: true }).decode(bytes);
  assert.ok(text.startsWith(head), 'what came before stays, mark and line ends included');
  assert.ok(text.endsWith(tail));
  assert.ok(text.slice(head.length, text.length - tail.length).startsWith('<w:font w:name="Fancy Face">'));

  // A root without content.
  const empty = await open({ body: para(run('here')), parts: [['word/fontTable.xml', encode(`﻿<w:fonts xmlns:w="${W}"/>\r\n`)]], fontTable: '' });
  paste(empty.doc, fragmentOf(source.doc, range(at(0, 0), at(0, 11))), at(0, 4));
  await writeDocx(empty.pkg, empty.doc);
  const made = new TextDecoder('utf-8', { ignoreBOM: true }).decode((await empty.pkg.readPart('/word/fontTable.xml'))!);
  assert.ok(made.startsWith(`﻿<w:fonts xmlns:w="${W}"><w:font w:name="Fancy Face">`) && made.endsWith('</w:fonts>\r\n'), made);
  assert.deepEqual([...(await reread(empty)).fonts.keys()], ['Fancy Face']);
});

test('review: a font table that is full takes no font more, and the file can be read again', async () => {
  const source = await open(SOURCE);
  const target = await open({ body: para(run('here')) });
  for (let i = 0; i < 20000; i++) target.doc.fonts.set('F' + i, { name: 'F' + i });
  paste(target.doc, fragmentOf(source.doc, range(at(0, 0), at(0, 11))), at(0, 4));
  assert.equal(target.doc.fonts.size, 20000);
  assert.ok(!target.doc.fonts.has('Fancy Face'));
});

test('review: a part named like a font table that the document does not relate to is left alone; the fonts get a part of their own', async () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const orphan = `<w:fonts xmlns:w="${W}"><w:font w:name="Orphan"/></w:fonts>`;
  const source = await open(SOURCE);
  const target = await open({ body: para(run('here')), parts: [['word/fontTable.xml', orphan]] });
  assert.equal(target.doc.fonts.size, 0, 'the fixture relates no font table');
  paste(target.doc, fragmentOf(source.doc, range(at(0, 0), at(0, 11))), at(0, 4));
  await writeDocx(target.pkg, target.doc);
  assert.equal(decoder.decode((await target.pkg.readPart('/word/fontTable.xml'))!), orphan);
  assert.ok(decoder.decode((await target.pkg.readPart('/word/fontTable2.xml'))!).includes('<w:font w:name="Fancy Face">'));
  assert.deepEqual([...(await reread(target)).fonts.keys()], ['Fancy Face']);
});

// ----------------------------------------------------------- second review

test('second review: the font of a symbol comes with the symbol', async () => {
  const symbol = '<w:r><w:sym w:font="Sign Face" w:char="F050"/></w:r>';
  const source = await open({ body: para(run('a') + symbol + run('b')), fontTable: font('Sign Face', '<w:charset w:val="02"/>') });
  const target = await open({ body: para(run('here')) });
  const fragment = fragmentOf(source.doc, range(at(0, 0), at(0, 3)));
  assert.deepEqual([...fragment.fonts!.keys()], ['Sign Face']);
  paste(target.doc, fragment, at(0, 4));
  assert.deepEqual(target.doc.fonts.get('Sign Face'), { name: 'Sign Face', charset: '02' });
  assert.deepEqual([...(await reread(target)).fonts.keys()], ['Sign Face']);
});

test('second review: a font named under a prefix of the other WordprocessingML namespace is a font the file has', async () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const table = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:fonts xmlns:w="${W}" xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:font s:name="Own Face"/></w:fonts>`;
  const target = await open({ body: para(run('here', 'Own Face')), parts: [['word/fontTable.xml', table]], fontTable: '' });
  assert.deepEqual([...target.doc.fonts.keys()], ['Own Face']);
  const history = new DocHistory(target.doc);
  history.applyEdit(insertFragment(target.doc, at(0, 4), fragmentOf(target.doc, range(at(0, 0), at(0, 2)))), undefined);
  history.seal();
  assert.equal(await written(target), table);
});

test('second review: a font table that is not UTF-8 is not written to; the other bytes of it are not lost for a font', async () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const text = `<?xml version="1.0" encoding="windows-1254"?><w:fonts xmlns:w="${W}"><w:font w:name="Owný"/></w:fonts>`;
  const bytes = Uint8Array.from(text, (c) => c.charCodeAt(0));
  const source = await open(SOURCE);
  const target = await open({ body: para(run('here')), parts: [['word/fontTable.xml', bytes]], fontTable: '' });
  paste(target.doc, fragmentOf(source.doc, range(at(0, 0), at(0, 11))), at(0, 4));
  const report = { parts: [] as string[], warnings: [] as string[], copiedBlocks: 0, writtenBlocks: 0, droppedCharacters: 0 };
  await writeDocx(target.pkg, target.doc, { report: report as never });
  assert.deepEqual(await target.pkg.readPart('/word/fontTable.xml'), bytes);
  assert.ok(report.warnings.some((warning) => warning.includes('not UTF-8')), report.warnings.join('; '));
});

test('second review: a font table of very many fonts is looked through once for a save', async () => {
  let table = '';
  for (let i = 0; i < 20000; i++) table += `<w:font w:name="F${i}"/>`;
  const target = await open({ body: para(run('here')), fontTable: table });
  const before = await target.pkg.readPart('/word/fontTable.xml');
  const history = new DocHistory(target.doc);
  history.applyEdit(insertFragment(target.doc, at(0, 4), fragmentOf(target.doc, range(at(0, 0), at(0, 2)))), undefined);
  history.seal();
  const started = Date.now();
  await writeDocx(target.pkg, target.doc);
  // A guard against the quadratic search (3.5 s here before), not a measurement.
  assert.ok(Date.now() - started < 2000, `saved in ${Date.now() - started} ms`);
  assert.deepEqual(await target.pkg.readPart('/word/fontTable.xml'), before);
});

test('second review: a full font table is still read again after a paste that could not add to it', async () => {
  let table = '';
  for (let i = 0; i < 20000; i++) table += `<w:font w:name="F${i}"/>`;
  const source = await open(SOURCE);
  const target = await open({ body: para(run('here')), fontTable: table });
  paste(target.doc, fragmentOf(source.doc, range(at(0, 0), at(0, 11))), at(0, 4));
  assert.equal(target.doc.fonts.size, 20000);
  assert.equal((await reread(target)).fonts.size, 20000);
});
