// What the review of the copy and paste of documents found, each as the test
// that failed before it was put right.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import { DocLayout } from '../src/layout/doc/layout';
import { paragraphText, type DocModel, type Paragraph, type Table } from '../src/model/doc/types';
import { OpcPackage } from '../src/opc';
import { DocHistory, deleteRange, fragmentOf, insertFragment, type DocFragment, type DocPosition, type DocRange } from '../src/ops/doc/index';
import { MAX_STORED_BYTES, clearStoredCopies, cutLoses, setCopyMedia, sliceOf, storeCopy, storedCopies } from '../src/ui/doc-clipboard';
import { parseXml, type XmlElement } from '../src/xml';
import { measurer } from './doc-ops-fixtures';
import { buildDocx, type DocxSpec } from './docx-fixtures';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const W15 = ' xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"';
const STYLES =
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>';
const NUMBERING = '<w:abstractNum w:abstractNumId="4"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="4"/></w:num>';
const SECT = '<w:sectPr><w:pgSz w:w="9000" w:h="9000"/></w:sectPr>';
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 2, 0, 0, 0, 1, 8, 2, 0, 0, 0]);

const r = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (content = '', pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${content.charAt(0) === '<' ? content : content ? r(content) : ''}</w:p>`;
const LIST = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>';
/** A part whose root declares WordprocessingML and relationships and nothing else, but for `extra`. */
const bare = (name: string, inner: string, extra = '') => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${name} xmlns:w="${W}" xmlns:r="${R}"${extra}>${inner}</w:${name}>`;
const picture = (rId: string, id: number) =>
  `<w:r><w:drawing><wp:inline><wp:extent cx="381000" cy="152400"/><wp:docPr id="${id}" name="P${id}"/>` +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="x"/><pic:cNvPicPr/></pic:nvPicPr>' +
  `<pic:blipFill><a:blip r:embed="${rId}"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

interface Opened {
  pkg: OpcPackage;
  doc: DocModel;
}

async function open(spec: DocxSpec | string): Promise<Opened> {
  const full: DocxSpec = typeof spec === 'string' ? { body: spec } : { ...spec };
  if (full.styles === undefined) full.styles = STYLES;
  if (full.body !== undefined && full.body.indexOf('<w:sectPr') < 0) full.body += SECT;
  const pkg = await OpcPackage.open(await buildDocx(full));
  return { pkg, doc: await readDocx(pkg) };
}

const at = (path: number[] | number, offset: number): DocPosition => ({ story: 'body', path: typeof path === 'number' ? [path] : path, offset });
const range = (a: DocPosition, b: DocPosition): DocRange => ({ anchor: a, focus: b });
const decoder = new TextDecoder();

function paste(doc: DocModel, fragment: DocFragment, where: DocPosition | DocRange, moved = false): DocHistory {
  const history = new DocHistory(doc);
  history.applyEdit(insertFragment(doc, where, fragment, moved ? { moved } : undefined), undefined);
  history.seal();
  return history;
}

/** Every prefix of an element or attribute name of the XML has a declaration in scope: the names of those that have none. */
function undeclared(xml: string): string[] {
  const out: string[] = [];
  const walk = (el: XmlElement, scope: Set<string>): void => {
    let own = scope;
    el.attrs.forEach((_value, name) => {
      if (name !== 'xmlns' && name.indexOf('xmlns:') !== 0) return;
      if (own === scope) own = new Set(scope);
      own.add(name === 'xmlns' ? '' : name.slice(6));
    });
    const check = (name: string): void => {
      const colon = name.indexOf(':');
      if (colon > 0 && name.indexOf('xmlns:') !== 0 && !own.has(name.slice(0, colon))) out.push(name);
    };
    check(el.name);
    el.attrs.forEach((_value, name) => check(name));
    for (const node of el.children) if (node.type === 'element') walk(node, own);
  };
  walk(parseXml(xml).root, new Set(['xml']));
  return out;
}

const info = (doc: DocModel): string[] =>
  doc.body.blocks.map((block) => {
    if (block.type !== 'paragraph') return block.type;
    const props = block.props;
    return `${paragraphText(block)}{${props.style || ''},${props.align || ''},${props.numbering ? props.numbering.numId : ''}}`;
  });

test('review 1: the definition of a style or a list that goes into another document declares the prefixes the styles part and the numbering part of that document do not', async () => {
  // A style with a w14 child. The document part of the destination declares w14; its styles part does not.
  const FANCY = '<w:style w:type="paragraph" w:styleId="Fancy"><w:name w:val="Fancy"/><w:rPr><w14:ligatures w14:val="all"/></w:rPr></w:style>';
  const source = await open({ body: p('styled', '<w:pStyle w:val="Fancy"/>') + p('x'), styles: STYLES + FANCY });
  const target = await open({ body: p('dest'), parts: [['word/styles.xml', bare('styles', STYLES)]] });
  paste(target.doc, fragmentOf(source.doc, range(at(0, 0), at(1, 0))), at(0, 0));
  await writeDocx(target.pkg, target.doc);
  let xml = decoder.decode(await target.pkg.readPart('/word/styles.xml'));
  assert.ok(xml.indexOf('w:styleId="Fancy"') > 0, xml);
  assert.ok(xml.indexOf('<w14:ligatures') > 0, xml);
  assert.deepEqual(undeclared(xml), []);
  assert.deepEqual(undeclared(decoder.decode(await target.pkg.readPart('/word/document.xml'))), []);
  let again = await readDocx(await OpcPackage.open(await target.pkg.save()));
  assert.equal(again.styles.styles.get('Fancy')!.run!.extras!.length, 1);

  // A list whose level has a w15 attribute and a w15 child; the numbering part of the source declares w15, nothing of
  // the destination does, and the document part of the source does not either.
  const LEVELS =
    '<w:abstractNum w:abstractNumId="4"><w:multiLevelType w:val="hybridMultilevel"/><w:lvl w:ilvl="0" w15:tentative="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:rPr><w15:x w15:val="1"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="4"/></w:num>';
  const OWN = '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
  for (const withPart of [true, false]) {
    const listed = await open({ numbering: '', document: bare('document', `<w:body>${p('item', LIST)}${p('x')}${SECT}</w:body>`), parts: [['word/numbering.xml', bare('numbering', LEVELS, W15)]] });
    const fragment = fragmentOf(listed.doc, range(at(0, 0), at(1, 0)));
    const spec: DocxSpec = { document: bare('document', `<w:body>${p('dest')}${SECT}</w:body>`) };
    if (withPart) {
      spec.numbering = '';
      spec.parts = [['word/numbering.xml', bare('numbering', OWN)]];
    }
    const other = await open(spec);
    paste(other.doc, fragment, at(0, 0));
    await writeDocx(other.pkg, other.doc);
    xml = decoder.decode(await other.pkg.readPart('/word/numbering.xml'));
    assert.ok(xml.indexOf('w15:tentative="1"') > 0 && xml.indexOf('<w15:x') > 0, xml);
    assert.deepEqual(undeclared(xml), [], xml);
    again = await readDocx(await OpcPackage.open(await other.pkg.save()));
    const numId = (again.body.blocks[0] as Paragraph).props.numbering!.numId!;
    assert.equal(again.numbering.abstracts.get(again.numbering.nums.get(numId)!.abstractId)!.levels.get(0)!.format, 'decimal');
  }
  // In the document they came from nothing is added to them.
  const same = await open({ body: p('styled', '<w:pStyle w:val="Fancy"/>') + p('x'), styles: STYLES + FANCY });
  paste(same.doc, fragmentOf(same.doc, range(at(0, 0), at(1, 0))), at(1, 1));
  assert.equal((same.doc.body.blocks[1] as Paragraph).props, (same.doc.body.blocks[0] as Paragraph).props);
});

test('review 2: a cut from one paragraph into another, pasted back, leaves both paragraphs formatted as they were; a copy pasted elsewhere leaves the paragraph it goes into its formatting', async () => {
  const BODY = p('abc', '<w:pStyle w:val="Heading1"/>') + p('def', LIST + '<w:jc w:val="center"/><w:keepNext/>') + p('tail') + p('other', '<w:jc w:val="right"/>');
  const opened = await open({ body: BODY, numbering: NUMBERING });
  const { doc } = opened;
  const start = info(doc);
  assert.deepEqual(start, ['abc{Heading1,,}', 'def{,center,7}', 'tail{,,}', 'other{,right,}']);
  const kept = (doc.body.blocks[1] as Paragraph).props;
  const selection = range(at(0, 1), at(1, 1));
  const slice = sliceOf(doc, selection.anchor, selection.focus)!;
  assert.equal(slice.fragment!.endProps, kept);
  assert.equal(cutLoses(doc, selection, slice), false);
  // The text alone does not hold the formatting the deletion takes: such a cut is refused.
  assert.equal(cutLoses(doc, selection, undefined), true);
  assert.equal(cutLoses(doc, selection, { paragraphs: slice.paragraphs }), true);
  assert.equal(cutLoses(doc, range(at(1, 0), at(1, 2)), undefined), false);

  const history = new DocHistory(doc);
  const removal = deleteRange(doc, selection);
  history.applyEdit(removal, undefined);
  history.seal();
  assert.deepEqual(info(doc), ['aef{Heading1,,}', 'tail{,,}', 'other{,right,}']);
  history.applyEdit(insertFragment(doc, removal.selection, slice.fragment!, { moved: true }), undefined);
  history.seal();
  assert.deepEqual(info(doc), start);
  // The very properties, with everything they hold.
  assert.equal((doc.body.blocks[1] as Paragraph).props, kept);
  assert.equal((doc.body.blocks[1] as Paragraph).props.keepNext, true);
  history.undo();
  // The cut pasted somewhere else: the text after the caret keeps the formatting of its paragraph.
  history.applyEdit(insertFragment(doc, at(2, 2), slice.fragment!, { moved: true }), undefined);
  assert.deepEqual(info(doc), ['aef{Heading1,,}', 'tail{,,}', 'otbc{Heading1,,}', 'dher{,right,}']);
  history.undo();
  history.undo();
  assert.deepEqual(info(doc), start);
  // A copy (nothing was cut) pasted into a paragraph formatted like its first one: the paragraph keeps its formatting too.
  history.applyEdit(insertFragment(doc, at(0, 3), slice.fragment!), undefined);
  assert.deepEqual(info(doc), ['abcbc{Heading1,,}', 'd{Heading1,,}', 'def{,center,7}', 'tail{,,}', 'other{,right,}']);
  history.undo();

  // An empty first paragraph: the deletion leaves the formatting of the second, and the paste gives the first its own.
  const empty = await open({ body: p('', '<w:pStyle w:val="Heading1"/>') + p('def', LIST) + p('tail'), numbering: NUMBERING });
  const second = range(at(0, 0), at(1, 2));
  const cut = sliceOf(empty.doc, second.anchor, second.focus)!;
  assert.equal(cutLoses(empty.doc, second, cut), false);
  const removed = deleteRange(empty.doc, second);
  const emptyHistory = new DocHistory(empty.doc);
  emptyHistory.applyEdit(removed, undefined);
  emptyHistory.seal();
  assert.deepEqual(info(empty.doc), ['f{,,7}', 'tail{,,}']);
  emptyHistory.applyEdit(insertFragment(empty.doc, removed.selection, cut.fragment!, { moved: true }), undefined);
  assert.deepEqual(info(empty.doc), ['{Heading1,,}', 'def{,,7}', 'tail{,,}']);
  // Saved and read again.
  await writeDocx(opened.pkg, opened.doc);
  assert.deepEqual(info(await readDocx(await OpcPackage.open(await opened.pkg.save()))), start);
});

test('review 3: a list whose level names a style that names the list comes into another document as one list', async () => {
  const LEVELS =
    '<w:abstractNum w:abstractNumId="4"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Clause"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="4"/></w:num>';
  const CLAUSE = '<w:style w:type="paragraph" w:styleId="Clause"><w:name w:val="Clause"/><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr></w:style>';
  const source = await open({ body: p('one', LIST) + p('two', '<w:pStyle w:val="Clause"/>') + p('x'), styles: STYLES + CLAUSE, numbering: LEVELS });
  const fragment = fragmentOf(source.doc, range(at(0, 0), at(2, 0)));
  const target = await open(p('dest'));
  const history = paste(target.doc, fragment, at(0, 0));
  const numbering = target.doc.numbering;
  assert.equal(numbering.abstracts.size, 1);
  assert.equal(numbering.nums.size, 1);
  const numId = (target.doc.body.blocks[0] as Paragraph).props.numbering!.numId!;
  assert.ok(numbering.nums.has(numId));
  assert.equal(target.doc.styles.styles.get('Clause')!.para!.numbering!.numId, numId);
  assert.equal(numbering.abstracts.get(numbering.nums.get(numId)!.abstractId)!.levels.get(0)!.paraStyle, 'Clause');
  await writeDocx(target.pkg, target.doc);
  const again = await readDocx(await OpcPackage.open(await target.pkg.save()));
  assert.deepEqual([again.numbering.abstracts.size, again.numbering.nums.size], [1, 1]);
  history.undo();
  assert.deepEqual([numbering.abstracts.size, numbering.nums.size, target.doc.styles.styles.has('Clause')], [0, 0, false]);
});

test('review 5: a copy that held more than the page keeps takes no more pictures', async () => {
  clearStoredCopies();
  const BODY = p(r('ab') + picture('rIdA', 1) + picture('rIdB', 2)) + p('x');
  const opened = await open({
    body: BODY,
    rels: [
      ['rIdA', 'image', 'media/image1.png'],
      ['rIdB', 'image', 'media/image2.png'],
    ],
    parts: [
      ['word/media/image1.png', PNG],
      ['word/media/image2.png', PNG],
    ],
  });
  const copy = storeCopy('ab', sliceOf(opened.doc, at(0, 0), at(1, 0))!);
  setCopyMedia(copy, '/word/media/image1.png', new Uint8Array(MAX_STORED_BYTES + 1));
  const media = copy.slice.fragment!.media;
  assert.equal(media.get('/word/media/image1.png')!.bytes, undefined);
  assert.equal(copy.full, true);
  // The picture that is read later is not added to what was cut down.
  setCopyMedia(copy, '/word/media/image2.png', new Uint8Array(1024));
  assert.equal(media.get('/word/media/image2.png')!.bytes, undefined);
  assert.ok(storedCopies().bytes <= MAX_STORED_BYTES);
  // It can still be pasted where the pictures are.
  paste(opened.doc, copy.slice.fragment!, at(1, 1));
  assert.equal((opened.doc.body.blocks[1] as Paragraph).children.length > 1, true);
  clearStoredCopies();
});

test('review 6b: a paste that adds styles and a list makes the layout start again, with counters that know them', async () => {
  const source = await open({ body: p('one', LIST + '<w:pStyle w:val="Heading1"/>') + p('two', LIST) + p('x'), numbering: NUMBERING });
  const fragment = fragmentOf(source.doc, range(at(0, 0), at(2, 0)));
  const target = await open({ body: p('dest'), styles: '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' });
  const layout = new DocLayout(target.doc, measurer, {});
  layout.layout();
  paste(target.doc, fragment, at(0, 0));
  // The layout does not follow such an edit: the one who holds it makes a new one (DocumentView.follow does).
  assert.equal(layout.edited(), undefined);
  const fresh = new DocLayout(target.doc, measurer, {});
  fresh.layout();
  const texts = fresh.pages[0].items.filter((item) => item.type === 'text').map((item) => (item as { text: string }).text);
  assert.ok(texts.indexOf('1.') >= 0 && texts.indexOf('2.') >= 0, JSON.stringify(texts));
});

test('review 6c: saved, the paste undone, saved again: no part, relationship or definition of the paste is left in the package', async () => {
  const source = await open({
    body: p(r('ab') + picture('rIdImg', 5) + '<w:hyperlink r:id="rIdL">' + r('link') + '</w:hyperlink>', LIST + '<w:pStyle w:val="Heading1"/>') + p('x'),
    numbering: NUMBERING,
    rels: [
      ['rIdImg', 'image', 'media/image1.png'],
      ['rIdL', 'hyperlink', 'https://example.com/', true],
    ],
    parts: [['word/media/image1.png', PNG]],
  });
  const fragment = fragmentOf(source.doc, range(at(0, 0), at(1, 0)));
  fragment.media.get('/word/media/image1.png')!.bytes = PNG;
  const target = await open({ body: p('dest'), styles: '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' });
  const names = target.pkg.partNames().sort();
  const before: { [part: string]: string } = {};
  for (const name of names) before[name] = decoder.decode(await target.pkg.readPart(name));
  const history = paste(target.doc, fragment, at(0, 0));
  await writeDocx(target.pkg, target.doc);
  const saved = await OpcPackage.open(await target.pkg.save());
  assert.ok(saved.hasPart('/word/media/image1.png'));
  assert.ok(saved.hasPart('/word/numbering.xml'));
  assert.ok(decoder.decode(await saved.readPart('/word/styles.xml')).indexOf('Heading1') > 0);
  assert.ok(decoder.decode(await saved.readPart('/word/_rels/document.xml.rels')).indexOf('https://example.com/') > 0);

  history.undo();
  await writeDocx(target.pkg, target.doc);
  const again = await OpcPackage.open(await target.pkg.save());
  assert.deepEqual(again.partNames().sort(), names);
  for (const name of names) assert.equal(decoder.decode(await again.readPart(name)), before[name], name);
  assert.equal(again.contentTypeOf('/word/media/image1.png'), undefined);
});

test('review 7: the attributes of tables, rows, cells and whole paragraphs go with them, but for those that name one element', async () => {
  const TABLE =
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
    '<w:tr w:rsidR="00AA11BB" w14:paraId="11223344" w14:textId="77777777"><w:trPr><w14:foo w14:val="1"/></w:trPr><w:tc w:id="cell7"><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>' +
    '<w:p w:rsidR="00CC22DD" w14:paraId="55667788"><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const BODY = p('a') + TABLE + p('mid').replace('<w:p>', '<w:p w:rsidP="00EE33FF" w14:paraId="0ABCDEF0">') + p('z');
  const opened = await open(BODY);
  const { doc } = opened;
  const selection = range(at(0, 1), at(3, 0));
  const slice = sliceOf(doc, selection.anchor, selection.focus)!;
  assert.equal(cutLoses(doc, selection, slice), false);
  const history = new DocHistory(doc);
  const removal = deleteRange(doc, selection);
  history.applyEdit(removal, undefined);
  history.seal();
  history.applyEdit(insertFragment(doc, removal.selection, slice.fragment!, { moved: true }), undefined);
  const table = doc.body.blocks[1] as Table;
  assert.deepEqual(table.rows[0].attrs, [['w:rsidR', '00AA11BB']]);
  assert.deepEqual(table.rows[0].cells[0].attrs, [['w:id', 'cell7']]);
  assert.deepEqual((table.rows[0].cells[0].blocks[0] as Paragraph).attrs, [['w:rsidR', '00CC22DD']]);
  assert.deepEqual((doc.body.blocks[2] as Paragraph).attrs, [['w:rsidP', '00EE33FF']]);
  await writeDocx(opened.pkg, opened.doc);
  const xml = decoder.decode(await opened.pkg.readPart('/word/document.xml'));
  assert.ok(xml.indexOf('<w:tr w:rsidR="00AA11BB">') > 0, xml);
  assert.equal(xml.indexOf('11223344'), -1);
  // A second paste of the same copy: no id of an element is in the document twice.
  history.applyEdit(insertFragment(doc, at(0, 0), slice.fragment!), undefined);
  const ids: string[] = [];
  JSON.stringify(doc.body.blocks, (key, value) => {
    if (Array.isArray(value) && value.length === 2 && typeof value[0] === 'string' && /paraId|textId/.test(value[0])) ids.push(value[1]);
    return value instanceof Map ? Array.from(value) : value;
  });
  assert.equal(new Set(ids).size, ids.length, ids.join());

  // Into a document that does not declare w14: the kept child of the row brings the declaration, and the part is well formed.
  const target = await open({ document: bare('document', `<w:body>${p('dest')}${SECT}</w:body>`) });
  paste(target.doc, fragmentOf((await open(BODY)).doc, range(at(0, 1), at(3, 0))), at(0, 0));
  await writeDocx(target.pkg, target.doc);
  const other = decoder.decode(await target.pkg.readPart('/word/document.xml'));
  assert.ok(other.indexOf('<w14:foo') > 0 && other.indexOf('w:rsidR="00AA11BB"') > 0, other);
  assert.deepEqual(undeclared(other), []);
});
