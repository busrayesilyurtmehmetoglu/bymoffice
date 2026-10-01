import assert from 'node:assert/strict';
import { test } from 'node:test';
import {
  findWrapper,
  paragraphText,
  rawOf,
  type Block,
  type DrawingItem,
  type FieldCharItem,
  type LegacyGraphicItem,
  type Paragraph,
  type Run,
  type RunItem,
  type Table,
} from '../src/model/doc/types';
import { serializeElement } from '../src/formats/docx/common';
import { readSpec } from './docx-fixtures';

const para = (blocks: Block[], index: number) => {
  const block = blocks[index];
  assert.equal(block.type, 'paragraph');
  return block as Paragraph;
};
const run = (paragraph: Paragraph, index: number) => {
  const inline = paragraph.children[index];
  assert.equal(inline.type, 'run');
  return inline as Run;
};
const types = (items: { type: string }[]) => items.map((item) => item.type);

test('docx: paragraphs, runs and the simple run items', async () => {
  const first =
    '<w:p w:rsidR="00AB12CD" w14:paraId="1A2B3C4D"><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr>' +
    '<w:r w:rsidRPr="00112233"><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Merhaba</w:t></w:r>' +
    '<w:r><w:t xml:space="preserve"> dünya </w:t><w:tab/><w:br/><w:br w:type="page"/><w:br w:type="column"/><w:cr/>' +
    '<w:noBreakHyphen/><w:softHyphen/><w:sym w:font="Wingdings" w:char="F0FC"/><w:t>  kırpılır\n </w:t>' +
    '<w:br w:type="textWrapping" w:clear="all"/><w:ptab w:alignment="right" w:relativeTo="margin" w:leader="dot"/>' +
    '<w:lastRenderedPageBreak/><w:pgNum/></w:r></w:p>';
  const doc = await readSpec({ body: first + '<w:p/><w:p><w:r><w:t>&lt;b&gt;A &amp; B&lt;/b&gt;</w:t></w:r></w:p>' });
  assert.equal(doc.body.blocks.length, 3);
  assert.equal(doc.strict, false);
  assert.deepEqual(doc.warnings, []);

  const p = para(doc.body.blocks, 0);
  assert.equal(rawOf(doc.body, p), first);
  assert.equal(p.props.style, 'Heading1');
  assert.equal(p.props.align, 'center');
  assert.deepEqual(p.attrs, [
    ['w:rsidR', '00AB12CD'],
    ['w14:paraId', '1A2B3C4D'],
  ]);

  const r0 = run(p, 0);
  assert.equal(r0.props.bold, true);
  assert.equal(r0.props.size, 28);
  assert.deepEqual(r0.attrs, [['w:rsidRPr', '00112233']]);
  assert.deepEqual(r0.content, [{ type: 'text', text: 'Merhaba' }]);

  const r1 = run(p, 1);
  assert.equal(r1.attrs, undefined);
  assert.deepEqual(r1.content, [
    { type: 'text', text: ' dünya ' },
    { type: 'tab' },
    { type: 'break', kind: 'line' },
    { type: 'break', kind: 'page' },
    { type: 'break', kind: 'column' },
    { type: 'break', kind: 'line', cr: true },
    { type: 'noBreakHyphen' },
    { type: 'softHyphen' },
    { type: 'symbol', char: '', font: 'Wingdings' },
    // Without xml:space="preserve" the outer whitespace does not count; what was written is kept beside it.
    { type: 'text', text: 'kırpılır', source: '  kırpılır\n ' },
    { type: 'break', kind: 'line', clear: 'all' },
    { type: 'positionalTab', alignment: 'right', relativeTo: 'margin', leader: 'dot' },
    { type: 'lastRenderedPageBreak' },
    { type: 'pageNumber' },
  ]);
  assert.equal(paragraphText(p), 'Merhaba dünya \t\n\n\n\n‑kırpılır\n');

  const empty = para(doc.body.blocks, 1);
  assert.deepEqual(empty.children, []);
  assert.equal(rawOf(doc.body, empty), '<w:p/>');
  // Markup in text stays text.
  assert.equal(paragraphText(para(doc.body.blocks, 2)), '<b>A & B</b>');
});

test('docx: equal properties share one object', async () => {
  const p = '<w:p><w:pPr><w:jc w:val="both"/><w:spacing w:after="120"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>x</w:t></w:r></w:p>';
  const doc = await readSpec({ body: p + p + '<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>y</w:t></w:r></w:p>' });
  const [a, b, c] = [0, 1, 2].map((i) => para(doc.body.blocks, i));
  assert.equal(a.props, b.props);
  assert.equal(run(a, 0).props, run(b, 0).props);
  assert.notEqual(a.props, c.props);
  // No properties at all is one shared empty object.
  assert.equal(run(c, 0).props, run(para((await readSpec({ body: '<w:p><w:r/></w:p>' })).body.blocks, 0), 0).props);
});

test('docx: wrappers around inline content are flattened onto the items', async () => {
  const doc = await readSpec({
    rels: [['rId7', 'hyperlink', 'https://example.com/a?b=1', true]],
    body:
      '<w:p>' +
      '<w:hyperlink r:id="rId7" w:tooltip="İpucu" w:history="1"><w:r><w:t>bağ</w:t></w:r>' +
      '<w:ins w:id="5" w:author="Ayşe" w:date="2024-01-02T03:04:05Z"><w:r><w:t>eklendi</w:t></w:r></w:ins></w:hyperlink>' +
      '<w:hyperlink w:anchor="Bolum1"/>' +
      '<w:del w:id="6" w:author="Ali" w:date="2024-01-03T00:00:00Z"><w:r><w:delText xml:space="preserve">silindi </w:delText></w:r></w:del>' +
      '<w:moveFrom w:id="7" w:author="Ali"><w:r><w:t>a</w:t></w:r></w:moveFrom><w:moveTo w:id="8" w:author="Ali"><w:r><w:t>b</w:t></w:r></w:moveTo>' +
      '<w:fldSimple w:instr=" PAGE " w:dirty="1"><w:fldData>AAA=</w:fldData><w:r><w:t>4</w:t></w:r></w:fldSimple>' +
      '<w:sdt><w:sdtPr><w:alias w:val="Ad"/><w:text/></w:sdtPr><w:sdtEndPr/><w:sdtContent><w:r><w:t>kontrol</w:t></w:r></w:sdtContent></w:sdt>' +
      '<w:smartTag w:uri="urn:x" w:element="place"><w:smartTagPr><w:attr w:name="a" w:val="b"/></w:smartTagPr><w:r><w:t>İzmir</w:t></w:r></w:smartTag>' +
      '<w:customXml w:uri="urn:y" w:element="item"><w:r><w:t>özel</w:t></w:r></w:customXml>' +
      '<w:bdo w:val="rtl"><w:r><w:t>ters</w:t></w:r></w:bdo>' +
      '<w:hyperlink r:id="rIdYok"><w:r><w:t>kırık</w:t></w:r></w:hyperlink>' +
      '</w:p>',
  });
  const p = para(doc.body.blocks, 0);
  assert.deepEqual(types(p.children), ['run', 'run', 'empty', 'run', 'run', 'run', 'run', 'run', 'run', 'run', 'run', 'run']);

  const link = findWrapper(p.children[0], 'hyperlink');
  assert.ok(link);
  assert.equal(link.rId, 'rId7');
  assert.equal(link.target, 'https://example.com/a?b=1');
  assert.equal(link.tooltip, 'İpucu');
  assert.deepEqual(link.attrs, [['w:history', '1']]);
  // The inserted run is inside the same hyperlink element.
  const inserted = findWrapper(p.children[1], 'ins');
  assert.ok(inserted);
  assert.deepEqual([inserted.id, inserted.author, inserted.date], [5, 'Ayşe', '2024-01-02T03:04:05Z']);
  assert.equal(inserted.parent, link);
  assert.equal(findWrapper(p.children[1], 'hyperlink'), link);

  // A wrapper without content is not lost.
  assert.equal(findWrapper(p.children[2], 'hyperlink')?.anchor, 'Bolum1');

  const deleted = findWrapper(p.children[3], 'del');
  assert.equal(deleted?.author, 'Ali');
  assert.deepEqual(run(p, 3).content, [{ type: 'deletedText', text: 'silindi ' }]);
  assert.equal(findWrapper(p.children[4], 'moveFrom')?.id, 7);
  assert.equal(findWrapper(p.children[5], 'moveTo')?.id, 8);

  const field = findWrapper(p.children[6], 'simpleField');
  assert.ok(field);
  assert.equal(field.instruction, ' PAGE ');
  assert.equal(field.dirty, true);
  assert.equal(field.extras?.length, 1);
  assert.equal(serializeElement(field.extras![0].el), '<w:fldData>AAA=</w:fldData>');
  assert.equal(field.extras![0].index, 0);

  const sdt = findWrapper(p.children[7], 'sdt');
  assert.ok(sdt?.pr && sdt.endPr);
  assert.equal(serializeElement(sdt.pr), '<w:sdtPr><w:alias w:val="Ad"/><w:text/></w:sdtPr>');

  const tag = findWrapper(p.children[8], 'smartTag');
  assert.deepEqual(tag?.attrs, [
    ['w:uri', 'urn:x'],
    ['w:element', 'place'],
  ]);
  assert.ok(tag?.pr);
  assert.equal(findWrapper(p.children[9], 'customXml')?.attrs?.length, 2);
  assert.equal(findWrapper(p.children[10], 'bdo')?.val, 'rtl');

  const broken = findWrapper(p.children[11], 'hyperlink');
  assert.equal(broken?.rId, 'rIdYok');
  assert.equal(broken?.target, undefined);
  assert.equal(doc.warnings.length, 1);
  // The text that was moved away ("a") is not shown; where it was moved to ("b") is.
  assert.equal(paragraphText(p), 'bağeklendib4kontrolİzmirözelterskırık');
});

test('docx: complex fields get their code and cached result, also across paragraphs and nested', async () => {
  const fld = (type: string, extra = '') => `<w:r><w:fldChar w:fldCharType="${type}"${extra}/></w:r>`;
  const doc = await readSpec({
    body:
      `<w:p><w:r><w:t xml:space="preserve">Sayfa </w:t></w:r>${fld('begin', ' w:dirty="true"')}<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>` +
      `<w:r><w:instrText>\\* MERGEFORMAT</w:instrText></w:r>${fld('separate')}<w:r><w:t>12</w:t></w:r>${fld('end')}</w:p>` +
      // Starts here, ends in the next paragraph; holds a nested field in its code.
      `<w:p>${fld('begin')}<w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>${fld('begin')}<w:r><w:instrText>DATE</w:instrText></w:r>` +
      `${fld('separate')}<w:r><w:t>bugün</w:t></w:r>${fld('end')}<w:r><w:instrText> = x</w:instrText></w:r>${fld('separate')}<w:r><w:t xml:space="preserve">sonuç </w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>devamı</w:t></w:r>${fld('end')}<w:r><w:t xml:space="preserve"> dışarıda</w:t></w:r></w:p>` +
      // A form field: begin carries data.
      '<w:p><w:r><w:fldChar w:fldCharType="begin" w:fldLock="1"><w:ffData><w:name w:val="Metin1"/></w:ffData></w:fldChar></w:r>' +
      `<w:r><w:instrText> FORMTEXT </w:instrText></w:r>${fld('end')}</w:p>`,
  });
  const p0 = para(doc.body.blocks, 0);
  const begin = run(p0, 1).content[0] as FieldCharItem;
  assert.equal(begin.charType, 'begin');
  assert.equal(begin.dirty, true);
  assert.deepEqual(begin.field, { instruction: ' PAGE \\* MERGEFORMAT', result: '12' });
  assert.deepEqual(
    p0.children.map((inline) => (inline as Run).fieldCode === true),
    [false, false, true, true, true, false, false],
  );
  assert.equal(paragraphText(p0), 'Sayfa 12');

  const p1 = para(doc.body.blocks, 1);
  const outer = run(p1, 0).content[0] as FieldCharItem;
  const inner = run(p1, 2).content[0] as FieldCharItem;
  assert.deepEqual(inner.field, { instruction: 'DATE', result: '' });
  // The nested result sits in the outer code, so it is not shown and not part of the outer result.
  assert.deepEqual(outer.field, { instruction: ' IF = x', result: 'sonuç devamı' });
  assert.equal(paragraphText(p1), 'sonuç ');
  assert.equal(paragraphText(para(doc.body.blocks, 2)), 'devamı dışarıda');

  const form = run(para(doc.body.blocks, 3), 0).content[0] as FieldCharItem;
  assert.equal(form.locked, true);
  assert.equal(form.data?.length, 1);
  assert.equal(serializeElement(form.data![0]), '<w:ffData><w:name w:val="Metin1"/></w:ffData>');
  assert.equal(form.field?.instruction, 'FORMTEXT');
});

test('docx: range marks, proofing marks, references and math', async () => {
  const doc = await readSpec({
    body:
      '<w:bookmarkStart w:id="0" w:name="_GoBack"/>' +
      '<w:p><w:bookmarkStart w:id="1" w:name="Bölüm" w:colFirst="0" w:colLast="1"/><w:commentRangeStart w:id="3"/>' +
      '<w:proofErr w:type="spellStart"/><w:r><w:t>metin</w:t></w:r><w:proofErr w:type="spellEnd"/>' +
      '<w:commentRangeEnd w:id="3"/><w:r><w:commentReference w:id="3"/></w:r><w:bookmarkEnd w:id="1"/>' +
      '<w:r><w:footnoteReference w:id="2"/><w:endnoteReference w:id="4" w:customMarkFollows="1"/><w:t>*</w:t></w:r>' +
      '<w:moveFromRangeStart w:id="9" w:name="move1" w:author="A" w:date="2024-01-01T00:00:00Z"/><w:moveFromRangeEnd w:id="9"/>' +
      '<w:permStart w:id="abc" w:edGrp="everyone"/><w:permEnd w:id="abc"/>' +
      '<m:oMath><m:r><m:t>x</m:t></m:r><m:sSup><m:e><m:r><m:t>y</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>' +
      '</w:p><w:bookmarkEnd w:id="0"/>',
  });
  assert.deepEqual(types(doc.body.blocks), ['mark', 'paragraph', 'mark']);
  const first = doc.body.blocks[0];
  assert.ok(first.type === 'mark');
  assert.deepEqual(first.mark, { type: 'bookmarkStart', id: '0', name: '_GoBack' });
  assert.equal(rawOf(doc.body, first), '<w:bookmarkStart w:id="0" w:name="_GoBack"/>');

  const p = para(doc.body.blocks, 1);
  assert.deepEqual(types(p.children), [
    'bookmarkStart',
    'commentRangeStart',
    'proofError',
    'run',
    'proofError',
    'commentRangeEnd',
    'run',
    'bookmarkEnd',
    'run',
    'moveFromRangeStart',
    'moveFromRangeEnd',
    'permStart',
    'permEnd',
    'math',
  ]);
  assert.deepEqual(p.children[0], {
    type: 'bookmarkStart',
    id: '1',
    name: 'Bölüm',
    attrs: [
      ['w:colFirst', '0'],
      ['w:colLast', '1'],
    ],
  });
  assert.deepEqual(p.children[2], { type: 'proofError', kind: 'spellStart' });
  assert.deepEqual(run(p, 6).content, [{ type: 'commentReference', id: 3 }]);
  assert.deepEqual(run(p, 8).content, [
    { type: 'footnoteReference', id: 2 },
    { type: 'endnoteReference', id: 4, customMark: true },
    { type: 'text', text: '*' },
  ]);
  assert.deepEqual(p.children[11], { type: 'permStart', id: 'abc', attrs: [['w:edGrp', 'everyone']] });
  const math = p.children[13];
  assert.ok(math.type === 'math');
  assert.equal(math.text, 'xy2');
  assert.equal(math.xml.name, 'm:oMath');
  assert.equal(paragraphText(p), 'metin*xy2');
  // Nothing here is "unknown".
  assert.equal(doc.unknown.size, 0);
});

const PICTURE = (rId: string) =>
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
  `<pic:blipFill><a:blip r:embed="${rId}"/><a:srcRect l="1000" t="2000" r="3000" b="4000"/><a:stretch/></pic:blipFill>` +
  '</pic:pic></a:graphicData></a:graphic>';

test('docx: drawings, text boxes, VML pictures, objects and alternate content', async () => {
  const shape =
    '<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">' +
    '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>-457200</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:align>top</wp:align></wp:positionV><wp:extent cx="1828800" cy="914400"/>' +
    '<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="2" name="Metin Kutusu 2" descr="Açıklama"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx><w:txbxContent>' +
    '<w:p><w:r><w:t>kutu</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>hücre</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
    '</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>';
  const pict =
    '<w:pict><v:shape id="s1" style="width:100pt;height:50pt"><v:imagedata r:id="rImg2" o:title="" xmlns:o="urn:schemas-microsoft-com:office:office"/>' +
    '<v:textbox><w:txbxContent><w:p><w:r><w:t>eski kutu</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict>';
  const doc = await readSpec({
    rels: [
      ['rImg1', 'image', 'media/image1.png'],
      ['rImg2', 'image', 'media/image2.emf'],
      ['rImg3', 'image', 'https://example.com/x.png', true],
    ],
    body:
      `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="952500" cy="476250"/><wp:docPr id="1" name="Resim 1"/>${PICTURE('rImg1')}</wp:inline></w:drawing></w:r>` +
      `<w:r><mc:AlternateContent><mc:Choice Requires="wps">${shape}</mc:Choice><mc:Fallback>${pict}</mc:Fallback></mc:AlternateContent></w:r>` +
      `<w:r><mc:AlternateContent><mc:Choice Requires="xyz">${shape}</mc:Choice><mc:Fallback>${pict}</mc:Fallback></mc:AlternateContent></w:r>` +
      '<w:r><w:object w:dxaOrig="1440" w:dyaOrig="1440"><v:shape id="o1"><v:imagedata r:id="rImg2"/></v:shape></w:object></w:r>' +
      '<w:r><w:drawing><wp:inline><wp:extent cx="-5" cy="99999999999"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<pic:pic><pic:blipFill><a:blip r:link="rImg3"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>' +
      '<w:r><w:drawing><wp:mystery/></w:drawing></w:r></w:p>',
  });
  const p = para(doc.body.blocks, 0);

  const picture = run(p, 0).content[0] as DrawingItem;
  assert.equal(picture.type, 'drawing');
  assert.equal(picture.placement, 'inline');
  assert.equal(picture.kind, 'picture');
  assert.deepEqual([picture.width, picture.height, picture.name], [952500, 476250, 'Resim 1']);
  assert.deepEqual(picture.image, { rId: 'rImg1', target: '/word/media/image1.png' });
  assert.deepEqual(picture.crop, { left: 1000, top: 2000, right: 3000, bottom: 4000 });
  assert.equal(picture.anchor, undefined);
  assert.equal(picture.xml.name, 'w:drawing');

  const chosen = run(p, 1).content[0];
  assert.ok(chosen.type === 'alternate');
  assert.equal(chosen.selected, 'choice');
  assert.equal(chosen.xml.name, 'mc:AlternateContent');
  const box = chosen.items[0] as DrawingItem;
  assert.equal(box.placement, 'anchor');
  assert.equal(box.kind, 'shape');
  assert.equal(box.description, 'Açıklama');
  assert.deepEqual(box.anchor, {
    distT: 0,
    distB: 0,
    distL: 114300,
    distR: 114300,
    relativeHeight: 251659264,
    behindDoc: true,
    locked: false,
    layoutInCell: true,
    allowOverlap: true,
    horizontal: { relativeFrom: 'column', offset: -457200 },
    vertical: { relativeFrom: 'paragraph', align: 'top' },
    wrap: 'square',
    wrapText: 'bothSides',
  });
  assert.equal(box.textBoxes?.length, 1);
  assert.deepEqual(types(box.textBoxes![0]), ['paragraph', 'table']);
  assert.equal(paragraphText(box.textBoxes![0][0] as Paragraph), 'kutu');

  // A requirement the reader does not know: the fallback is taken.
  const fallback = run(p, 2).content[0];
  assert.ok(fallback.type === 'alternate');
  assert.equal(fallback.selected, 'fallback');
  const legacy = fallback.items[0] as LegacyGraphicItem;
  assert.equal(legacy.type, 'pict');
  assert.deepEqual(legacy.image, { rId: 'rImg2', target: '/word/media/image2.emf' });
  assert.equal(paragraphText(legacy.textBoxes![0][0] as Paragraph), 'eski kutu');

  const object = run(p, 3).content[0] as LegacyGraphicItem;
  assert.equal(object.type, 'object');
  assert.equal(object.image?.rId, 'rImg2');
  assert.equal(object.xml.attrs.get('w:dxaOrig'), '1440');

  // Sizes outside the valid range are not taken; the linked image is external.
  const linked = run(p, 4).content[0] as DrawingItem;
  assert.equal(linked.width, undefined);
  assert.equal(linked.height, undefined);
  assert.deepEqual(linked.image, { rId: 'rImg3', linked: true, target: 'https://example.com/x.png', external: true });

  const mystery = run(p, 5).content[0];
  assert.ok(mystery.type === 'unknown');
  assert.equal(serializeElement(mystery.xml), '<w:drawing><wp:mystery/></w:drawing>');
});

test('docx: tables with properties, merged cells, nesting and row-level wrappers', async () => {
  const doc = await readSpec({
    body:
      '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/><w:tblInd w:w="-108" w:type="dxa"/>' +
      '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="dashed" w:sz="8" w:color="FF0000"/></w:tblBorders>' +
      '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>' +
      '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' +
      '<w:tr w:rsidR="00AA"><w:trPr><w:cantSplit/><w:trHeight w:val="567" w:hRule="exact"/><w:tblHeader/></w:trPr>' +
      '<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/><w:vAlign w:val="center"/>' +
      '<w:tcBorders><w:bottom w:val="double" w:sz="6"/><w:tl2br w:val="single"/></w:tcBorders><w:tcMar><w:top w:w="57" w:type="dxa"/></w:tcMar><w:noWrap/></w:tcPr>' +
      '<w:p><w:r><w:t>Başlık</w:t></w:r></w:p></w:tc>' +
      '<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr><w:p/></w:tc></w:tr>' +
      '<w:bookmarkStart w:id="4" w:name="satirlar"/>' +
      '<w:sdt><w:sdtPr><w:id w:val="77"/></w:sdtPr><w:sdtContent>' +
      '<w:tr><w:tblPrEx><w:jc w:val="right"/></w:tblPrEx><w:trPr><w:gridBefore w:val="1"/><w:wBefore w:w="3000" w:type="dxa"/></w:trPr>' +
      '<w:tc><w:p><w:r><w:t>dış</w:t></w:r></w:p><w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>iç</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:tc>' +
      '<w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr></w:sdtContent></w:sdt></w:tbl><w:p/>',
  });
  const table = doc.body.blocks[0] as Table;
  assert.equal(table.type, 'table');
  const tableSource = rawOf(doc.body, table);
  assert.ok(tableSource && tableSource.startsWith('<w:tbl>') && tableSource.endsWith('</w:tbl>'));
  assert.deepEqual(table.grid, [3000, 3000, 3000]);
  assert.deepEqual(table.props, {
    style: 'TableGrid',
    width: { type: 'pct', value: 5000 },
    align: 'center',
    indent: { type: 'dxa', value: -108 },
    borders: {
      top: { style: 'single', size: 4, space: 0, color: { auto: true } },
      insideH: { style: 'dashed', size: 8, color: { rgb: 'FF0000' } },
    },
    layout: 'fixed',
    cellMargins: { left: { type: 'dxa', value: 108 }, right: { type: 'dxa', value: 108 } },
    // Both spellings were written: the bit field and an attribute for each flag.
    look: { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true, val: 0x04a0, stated: 0x07e0 },
  });
  assert.equal(table.rows.length, 2);

  const [head, second] = table.rows;
  assert.deepEqual(head.props, { cantSplit: true, height: 567, heightRule: 'exact', header: true });
  assert.deepEqual(head.attrs, [['w:rsidR', '00AA']]);
  assert.deepEqual(head.cells[0].props, {
    width: { type: 'dxa', value: 6000 },
    gridSpan: 2,
    shading: { pattern: 'clear', color: { auto: true }, fill: { rgb: 'D9D9D9' } },
    vAlign: 'center',
    borders: { bottom: { style: 'double', size: 6 }, tl2br: { style: 'single' } },
    margins: { top: { type: 'dxa', value: 57 } },
    noWrap: true,
  });
  assert.equal(head.cells[1].props.vMerge, 'restart');

  // The row inside the content control is a row of the table, marked with its wrapper.
  assert.equal(second.wrap?.kind, 'sdt');
  assert.equal(head.wrap, undefined);
  assert.deepEqual(second.props, { gridBefore: 1, widthBefore: { type: 'dxa', value: 3000 } });
  assert.deepEqual(second.tableProps, { align: 'right' });
  assert.equal(second.cells[1].props.vMerge, 'continue');
  assert.deepEqual(types(second.cells[0].blocks), ['paragraph', 'table', 'paragraph']);
  const innerTable = second.cells[0].blocks[1] as Table;
  // Only blocks directly in the story say where they are in the source.
  assert.equal(rawOf(doc.body, innerTable), undefined);
  assert.equal(paragraphText(innerTable.rows[0].cells[0].blocks[0] as Paragraph), 'iç');

  // The bookmark between the rows is kept with its position among the children of the table.
  assert.equal(table.extras?.length, 1);
  assert.equal(table.extras![0].index, 3);
  assert.equal(serializeElement(table.extras![0].el), '<w:bookmarkStart w:id="4" w:name="satirlar"/>');
});

test('docx: sections, headers and footers', async () => {
  const doc = await readSpec({
    headers: { rIdH1: '<w:p><w:r><w:t>Üst</w:t></w:r></w:p>', rIdH2: '<w:p><w:r><w:t>İlk sayfa</w:t></w:r></w:p>' },
    footers: {
      rIdF1: '<w:p><w:r><w:t>Sayfa </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
    },
    body:
      '<w:p><w:pPr><w:sectPr w:rsidR="00B1"><w:type w:val="continuous"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape" w:code="9"/>' +
      '<w:cols w:num="2" w:space="708" w:equalWidth="0" w:sep="1"><w:col w:w="4000" w:space="500"/><w:col w:w="6000"/></w:cols></w:sectPr></w:pPr></w:p>' +
      '<w:p><w:r><w:t>son bölüm</w:t></w:r></w:p>' +
      '<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:headerReference w:type="first" r:id="rIdH2"/><w:footerReference w:type="default" r:id="rIdF1"/>' +
      '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="-1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>' +
      '<w:pgBorders w:offsetFrom="page"><w:top w:val="single" w:sz="4"/></w:pgBorders>' +
      '<w:pgNumType w:fmt="lowerRoman" w:start="3"/><w:cols w:space="708"/><w:titlePg/><w:vAlign w:val="center"/><w:docGrid w:type="lines" w:linePitch="360"/></w:sectPr>',
  });
  assert.equal(doc.body.blocks.length, 2);
  assert.deepEqual(para(doc.body.blocks, 0).props.section, {
    attrs: [['w:rsidR', '00B1']],
    type: 'continuous',
    page: { width: 16838, height: 11906, orientation: 'landscape', code: 9 },
    columns: { count: 2, space: 708, equalWidth: false, separator: true, columns: [{ width: 4000, space: 500 }, { width: 6000 }] },
  });

  const last = doc.body.section;
  assert.ok(last);
  assert.ok(rawOf(doc.body, doc.body.sectionSpan)?.startsWith('<w:sectPr>'));
  assert.equal(doc.body.sectionIndex, undefined);
  assert.deepEqual(last.headers, { default: 'rIdH1', first: 'rIdH2' });
  assert.deepEqual(last.footers, { default: 'rIdF1' });
  assert.deepEqual(last.page, { width: 11906, height: 16838 });
  assert.deepEqual(last.margins, { top: 1417, bottom: -1417, right: 1417, left: 1417, header: 708, footer: 708, gutter: 0 });
  assert.deepEqual(last.pageNumbering, { format: 'lowerRoman', start: 3 });
  assert.deepEqual(last.columns, { space: 708 });
  assert.equal(last.titlePage, true);
  assert.equal(last.vAlign, 'center');
  assert.deepEqual(last.docGrid, { type: 'lines', linePitch: 360 });
  assert.deepEqual(last.pageBorders, { offsetFrom: 'page', top: { style: 'single', size: 4 } });
  assert.equal(last.extras, undefined);

  assert.deepEqual([...doc.headers.keys()], ['rIdH1', 'rIdH2']);
  assert.equal(paragraphText(doc.headers.get('rIdH2')!.blocks[0] as Paragraph), 'İlk sayfa');
  const footer = doc.footers.get('rIdF1')!;
  assert.equal(footer.part, '/word/footer1.xml');
  assert.equal(footer.rootName, 'w:ftr');
  const begin = run(footer.blocks[0] as Paragraph, 1).content[0] as FieldCharItem;
  assert.deepEqual(begin.field, { instruction: 'PAGE', result: '' });
});

test('docx: block-level wrappers are flattened and keep their source text', async () => {
  const sdt =
    '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr><w:sdtContent>' +
    '<w:p><w:r><w:t>bir</w:t></w:r></w:p><w:sdt><w:sdtContent><w:p><w:r><w:t>iki</w:t></w:r></w:p></w:sdtContent></w:sdt></w:sdtContent></w:sdt>';
  const doc = await readSpec({
    body:
      sdt +
      '<w:sdt><w:sdtPr/><w:sdtContent/></w:sdt>' +
      '<w:customXml w:uri="urn:z" w:element="kayit"><w:customXmlPr><w:attr w:name="n" w:val="v"/></w:customXmlPr><w:p><w:r><w:t>üç</w:t></w:r></w:p></w:customXml>' +
      '<mc:AlternateContent><mc:Choice Requires="w14"><w:p><w:r><w:t>yeni</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>eski</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent>',
  });
  assert.deepEqual(types(doc.body.blocks), ['paragraph', 'paragraph', 'empty', 'paragraph', 'paragraph']);
  const [one, two, none, three, alt] = doc.body.blocks;
  assert.equal(one.wrap?.kind, 'sdt');
  assert.equal(rawOf(doc.body, one.wrap), sdt);
  assert.equal(rawOf(doc.body, one), undefined);
  assert.equal(two.wrap?.kind, 'sdt');
  assert.equal(two.wrap?.parent, one.wrap);
  assert.notEqual(two.wrap, one.wrap);
  assert.equal(none.wrap?.kind, 'sdt');
  assert.ok(three.wrap?.kind === 'customXml');
  assert.ok(three.wrap.pr);
  assert.ok(alt.wrap?.kind === 'alternate');
  assert.equal(alt.wrap.selected, 'choice');
  assert.equal(paragraphText(alt as Paragraph), 'yeni');
});

test('docx: footnotes, endnotes and comments', async () => {
  const doc = await readSpec({
    footnotes:
      '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Dipnot</w:t></w:r></w:p></w:footnote>' +
      '<w:footnote w:id="2"><w:p/></w:footnote><w:footnote><w:p/></w:footnote><w:footnote w:id="7"/>',
    endnotes: '<w:endnote w:id="4" w:type="normal"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t>Sonnot</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:endnote>',
    comments:
      '<w:comment w:id="3" w:author="Zeynep Kaya" w:date="2024-05-06T07:08:09Z" w:initials="ZK" w14:x="1"><w:p><w:r><w:annotationRef/></w:r><w:r><w:t>Yorum</w:t></w:r></w:p></w:comment>',
    body: '<w:p/>',
  });
  assert.ok(doc.footnotes && doc.endnotes && doc.comments);
  assert.deepEqual([...doc.footnotes.items.keys()], [-1, 0, 2, 7]);
  assert.equal(doc.footnotes.items.get(-1)!.kind, 'separator');
  assert.deepEqual(run(doc.footnotes.items.get(-1)!.blocks[0] as Paragraph, 0).content, [{ type: 'separator' }]);
  const note = doc.footnotes.items.get(2)!;
  assert.equal(note.kind, undefined);
  assert.deepEqual(run(note.blocks[0] as Paragraph, 0).content, [{ type: 'footnoteRef' }]);
  assert.equal(paragraphText(note.blocks[0] as Paragraph), ' Dipnot');
  assert.deepEqual(doc.footnotes.items.get(7)!.blocks, []);
  // The repeated id and the note without an id were skipped and reported.
  assert.equal(doc.warnings.length, 2);

  const endnote = doc.endnotes.items.get(4)!;
  assert.equal(endnote.kind, undefined);
  assert.deepEqual(types(endnote.blocks), ['paragraph', 'table']);
  assert.equal(doc.endnotes.part, '/word/endnotes.xml');

  const comment = doc.comments.items.get(3)!;
  assert.deepEqual([comment.author, comment.date, comment.initials], ['Zeynep Kaya', '2024-05-06T07:08:09Z', 'ZK']);
  assert.deepEqual(comment.attrs, [['w14:x', '1']]);
  assert.equal(paragraphText(comment.blocks[0] as Paragraph), 'Yorum');
});

test('docx: what the reader does not understand is kept in place', async () => {
  const doc = await readSpec({
    document:
      '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:x="urn:x">' +
      '<w:background w:color="FFFFFF"/><w:body>' +
      '<w:altChunk r:id="rId9"/>' +
      '<w:p><w:pPr><w:keepNext/><w:divId w:val="7"/><w:spacing w:before="120" w:beforeFoo="1"/><w:jc w:val="left"/>' +
      '<w:ind w:left="99999999"/><w:outlineLvl w:val="12"/><w:pPrChange w:id="1" w:author="A"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr>' +
      '<w:subDoc r:id="rId3"/>' +
      '<w:r><w:rPr><w:b/><w14:ligatures w14:val="all"/><w:sz w:val="abc"/><w:color w:val="GGGGGG"/><w:u w:val="single" x:extra="1"/><w:b w:val="0"/>' +
      '<w:rPrChange w:id="2" w:author="A"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr>' +
      '<w:t>a</w:t><w:ruby><w:rt><w:r><w:t>furi</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>base</w:t></w:r></w:rubyBase></w:ruby><x:custom a="1">metin</x:custom>' +
      '<w:t w:foo="1">b</w:t><w:br w:type="strange"/><w:tab w:x="1"/></w:r></w:p>' +
      '<w:tbl><w:tblPr><w:tblStyleMystery w:val="1"/><w:tblW w:w="abc" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="100"/><w:tblGridChange w:id="1"/></w:tblGrid>' +
      '<w:tr><w:trPr><w:divId w:val="1"/></w:trPr><w:tc><w:tcPr><w:cellIns w:id="3" w:author="A"/><w:gridSpan w:val="0"/></w:tcPr><w:p/></w:tc><w:mystery/></w:tr></w:tbl>' +
      '</w:body></w:document>',
  });
  assert.equal(doc.body.extras?.length, 1);
  assert.equal(doc.body.extras![0].index, 0);
  assert.equal(serializeElement(doc.body.extras![0].el), '<w:background w:color="FFFFFF"/>');

  const chunk = doc.body.blocks[0];
  assert.ok(chunk.type === 'unknown');
  assert.equal(serializeElement(chunk.xml), '<w:altChunk r:id="rId9"/>');
  assert.equal(rawOf(doc.body, chunk), '<w:altChunk r:id="rId9"/>');

  const p = para(doc.body.blocks, 1);
  assert.equal(p.props.keepNext, true);
  assert.equal(p.props.align, 'left');
  // The valid part of a partly understood element is read; the element is kept as a whole.
  assert.deepEqual(p.props.spacing, { before: 120 });
  assert.deepEqual(p.props.indent, {});
  assert.equal(p.props.outlineLevel, undefined);
  assert.equal(p.props.change?.name, 'w:pPrChange');
  assert.deepEqual(
    p.props.extras!.map((extra) => [extra.index, extra.el.name, extra.partial === true]),
    [
      [1, 'w:divId', false],
      [2, 'w:spacing', true],
      [4, 'w:ind', true],
      [5, 'w:outlineLvl', false],
    ],
  );

  const sub = p.children[0];
  assert.ok(sub.type === 'unknown');
  assert.equal(sub.xml.name, 'w:subDoc');

  const r = run(p, 1);
  assert.equal(r.props.bold, true);
  assert.equal(r.props.size, undefined);
  assert.equal(r.props.color, undefined);
  assert.deepEqual(r.props.underline, { style: 'single' });
  assert.equal(r.props.change?.name, 'w:rPrChange');
  assert.deepEqual(
    r.props.extras!.map((extra) => [extra.index, extra.el.name, extra.partial === true]),
    [
      [1, 'w14:ligatures', false],
      [2, 'w:sz', false],
      [3, 'w:color', false],
      [4, 'w:u', true],
      // A property stated twice: the first is read, the second is kept as it is.
      [5, 'w:b', false],
    ],
  );
  assert.deepEqual(types(r.content), ['text', 'unknown', 'unknown', 'unknown', 'unknown', 'unknown']);
  const ruby = r.content[1] as Extract<RunItem, { type: 'unknown' }>;
  assert.equal(ruby.xml.name, 'w:ruby');
  assert.equal(serializeElement((r.content[2] as typeof ruby).xml), '<x:custom a="1">metin</x:custom>');
  assert.equal(serializeElement((r.content[3] as typeof ruby).xml), '<w:t w:foo="1">b</w:t>');

  const table = doc.body.blocks[2] as Table;
  assert.deepEqual(table.props.extras!.map((extra) => extra.el.name), ['w:tblStyleMystery', 'w:tblW']);
  assert.equal(table.props.width, undefined);
  assert.deepEqual(table.grid, [100]);
  assert.deepEqual(table.extras!.map((extra) => [extra.el.name, extra.partial]), [['w:tblGrid', true]]);
  assert.deepEqual(table.rows[0].props.extras!.map((extra) => extra.el.name), ['w:divId']);
  assert.deepEqual(table.rows[0].extras!.map((extra) => [extra.index, extra.el.name]), [[2, 'w:mystery']]);
  assert.deepEqual(table.rows[0].cells[0].props.extras!.map((extra) => extra.el.name), ['w:cellIns', 'w:gridSpan']);

  // The tally of what was kept names each of them.
  assert.equal(doc.unknown.get('rPr/w14:ligatures'), 1);
  assert.equal(doc.unknown.get('pPr/w:spacing (partial)'), 1);
  assert.equal(doc.unknown.get('body/w:altChunk'), 1);
  assert.equal(doc.unknown.get('r/w:ruby'), 1);
  assert.equal(doc.unknown.get('document/w:background'), 1);
});

/** Blocks without their places in the source, for comparing two readings. */
function plain(value: unknown): unknown {
  return JSON.parse(JSON.stringify(value, (key, item) => (key === 'srcStart' || key === 'srcEnd' ? undefined : item)));
}

test('docx: a strict document reads the same as a transitional one', async () => {
  const body =
    '<w:p><w:pPr><w:pStyle w:val="Baslik"/><w:ind w:left="720" w:hanging="360"/><w:tabs><w:tab w:val="left" w:pos="1440"/></w:tabs></w:pPr>' +
    '<w:hyperlink r:id="rId5"><w:r><w:rPr><w:b/><w:color w:val="FF0000" w:themeColor="accent1" w:themeTint="99"/></w:rPr><w:t>bağlantı</w:t></w:r></w:hyperlink></w:p>' +
    '<w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>';
  const spec = {
    body,
    headers: { rIdH: '<w:p><w:r><w:t>üst</w:t></w:r></w:p>' },
    styles: '<w:style w:type="paragraph" w:styleId="Baslik"><w:name w:val="Başlık"/><w:rPr><w:sz w:val="32"/></w:rPr></w:style>',
    rels: [['rId5', 'hyperlink', 'https://example.com', true]] as [string, string, string, boolean][],
  };
  const transitional = await readSpec(spec);
  const strict = await readSpec({ ...spec, strict: true });
  assert.equal(transitional.strict, false);
  assert.equal(strict.strict, true);
  assert.deepEqual(plain(strict.body.blocks), plain(transitional.body.blocks));
  assert.deepEqual(strict.body.section, transitional.body.section);
  assert.deepEqual(plain(strict.headers.get('rIdH')!.blocks), plain(transitional.headers.get('rIdH')!.blocks));
  assert.deepEqual([...strict.styles.styles], [...transitional.styles.styles]);
  assert.equal(findWrapper(para(strict.body.blocks, 0).children[0], 'hyperlink')?.target, 'https://example.com');
  assert.deepEqual(run(para(strict.body.blocks, 0), 0).props.color, { rgb: 'FF0000', theme: 'accent1', tint: 0x99 });

  // Strict documents may give measurements with units.
  const units = await readSpec({
    strict: true,
    body:
      '<w:p><w:pPr><w:ind w:start="1in" w:end="2.54cm" w:firstLine="12pt"/><w:spacing w:before="10mm" w:line="1pc"/><w:jc w:val="start"/></w:pPr>' +
      '<w:r><w:rPr><w:sz w:val="14pt"/><w:w w:val="150%"/></w:rPr><w:t>x</w:t></w:r></w:p>' +
      '<w:tbl><w:tblPr><w:tblW w:w="50%" w:type="pct"/></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>',
  });
  const p = para(units.body.blocks, 0);
  assert.deepEqual(p.props.indent, { left: 1440, right: 1440, firstLine: 240, logical: true });
  assert.deepEqual(p.props.spacing, { before: 567, line: 240 });
  assert.equal(p.props.align, 'start');
  assert.equal(run(p, 0).props.size, 28);
  assert.equal(run(p, 0).props.scale, 150);
  assert.deepEqual((units.body.blocks[1] as Table).props.width, { type: 'pct', value: 2500 });
  assert.equal(units.unknown.size, 0);
});

test('docx: prefixes other than the usual ones, and a default namespace, are followed', async () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const prefixed = await readSpec({
    document:
      `<x:document xmlns:x="${W}" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x:body>` +
      '<x:p><x:pPr><x:jc x:val="right"/></x:pPr><x:hyperlink rel:id="rId5"><x:r><x:rPr><x:i/></x:rPr><x:t>metin</x:t></x:r></x:hyperlink></x:p>' +
      '<x:sectPr><x:footerReference x:type="even" rel:id="rIdF"/></x:sectPr></x:body></x:document>',
    rels: [['rId5', 'hyperlink', 'https://example.com', true]],
  });
  const p = para(prefixed.body.blocks, 0);
  assert.equal(p.props.align, 'right');
  assert.equal(run(p, 0).props.italic, true);
  assert.equal(findWrapper(p.children[0], 'hyperlink')?.target, 'https://example.com');
  assert.deepEqual(prefixed.body.section?.footers, { even: 'rIdF' });
  assert.equal(prefixed.body.rootName, 'x:document');

  const unprefixed = await readSpec({
    document: `<document xmlns="${W}" xmlns:w="${W}"><body><p><pPr><jc w:val="center"/></pPr><r><t>düz</t></r></p></body></document>`,
  });
  assert.equal(para(unprefixed.body.blocks, 0).props.align, 'center');
  assert.equal(paragraphText(para(unprefixed.body.blocks, 0)), 'düz');
});

test('docx: paragraph and run properties are read in full', async () => {
  const doc = await readSpec({
    body:
      '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:keepNext/><w:keepLines w:val="0"/><w:pageBreakBefore w:val="true"/><w:widowControl w:val="off"/>' +
      '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr>' +
      '<w:pBdr><w:top w:val="single" w:sz="6" w:space="1" w:color="auto"/><w:left w:val="nil"/><w:between w:val="dotted" w:sz="4" w:themeColor="accent2" w:themeShade="BF"/></w:pBdr>' +
      '<w:shd w:val="pct10" w:color="000000" w:fill="FFFF00" w:themeFill="background1" w:themeFillTint="33"/>' +
      '<w:tabs><w:tab w:val="clear" w:pos="720"/><w:tab w:val="right" w:leader="dot" w:pos="9062"/></w:tabs>' +
      '<w:suppressAutoHyphens/><w:bidi/><w:spacing w:before="240" w:after="0" w:line="360" w:lineRule="auto" w:beforeAutospacing="1"/>' +
      '<w:ind w:left="1440" w:right="-200" w:hanging="360" w:leftChars="100"/><w:contextualSpacing/><w:mirrorIndents/><w:jc w:val="both"/>' +
      '<w:textAlignment w:val="center"/><w:outlineLvl w:val="2"/><w:cnfStyle w:val="100000000000" w:firstRow="1"/>' +
      '<w:rPr><w:b/><w:del w:id="9" w:author="Can" w:date="2024-02-02T00:00:00Z"/></w:rPr></w:pPr>' +
      '<w:r><w:rPr><w:rStyle w:val="Strong"/><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="MS Mincho" w:cs="Times New Roman" w:asciiTheme="minorHAnsi" w:cstheme="minorBidi" w:hint="eastAsia"/>' +
      '<w:b/><w:bCs w:val="0"/><w:i/><w:iCs/><w:caps/><w:smallCaps w:val="false"/><w:strike/><w:dstrike/><w:outline/><w:shadow/><w:emboss/><w:imprint/>' +
      '<w:noProof/><w:vanish/><w:webHidden/><w:color w:val="auto"/><w:spacing w:val="-20"/><w:w w:val="90"/><w:kern w:val="28"/><w:position w:val="-6"/>' +
      '<w:sz w:val="21"/><w:szCs w:val="22"/><w:highlight w:val="yellow"/><w:u w:val="wave" w:color="00B050"/><w:effect w:val="none"/>' +
      '<w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto" w:frame="1"/><w:shd w:val="clear" w:fill="auto"/><w:vertAlign w:val="superscript"/>' +
      '<w:rtl/><w:cs/><w:em w:val="dot"/><w:lang w:val="tr-TR" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr><w:t>x</w:t></w:r></w:p>',
  });
  const p = para(doc.body.blocks, 0);
  assert.deepEqual(p.props, {
    style: 'ListParagraph',
    keepNext: true,
    keepLines: false,
    pageBreakBefore: true,
    widowControl: false,
    numbering: { level: 1, numId: 3 },
    borders: {
      top: { style: 'single', size: 6, space: 1, color: { auto: true } },
      left: { style: 'nil' },
      between: { style: 'dotted', size: 4, color: { theme: 'accent2', shade: 0xbf } },
    },
    shading: { pattern: 'pct10', fill: { rgb: 'FFFF00', theme: 'background1', tint: 0x33 }, color: { rgb: '000000' } },
    tabs: [
      { pos: 720, align: 'clear' },
      { pos: 9062, align: 'right', leader: 'dot' },
    ],
    suppressAutoHyphens: true,
    bidi: true,
    spacing: { before: 240, after: 0, line: 360, lineRule: 'auto', beforeAuto: true },
    indent: { left: 1440, right: -200, hanging: 360, leftChars: 100 },
    contextualSpacing: true,
    mirrorIndents: true,
    align: 'both',
    textAlignment: 'center',
    outlineLevel: 2,
    cnfStyle: '100000000000',
    markProps: { bold: true, deleted: { id: 9, author: 'Can', date: '2024-02-02T00:00:00Z' } },
  });
  assert.deepEqual(run(p, 0).props, {
    style: 'Strong',
    fonts: {
      ascii: 'Arial',
      hAnsi: 'Arial',
      eastAsia: 'MS Mincho',
      cs: 'Times New Roman',
      asciiTheme: 'minorHAnsi',
      csTheme: 'minorBidi',
      hint: 'eastAsia',
    },
    bold: true,
    boldCs: false,
    italic: true,
    italicCs: true,
    caps: true,
    smallCaps: false,
    strike: true,
    doubleStrike: true,
    outline: true,
    shadow: true,
    emboss: true,
    imprint: true,
    noProof: true,
    vanish: true,
    webHidden: true,
    color: { auto: true },
    spacing: -20,
    scale: 90,
    kerning: 28,
    position: -6,
    size: 21,
    sizeCs: 22,
    highlight: 'yellow',
    underline: { style: 'wave', color: { rgb: '00B050' } },
    effect: 'none',
    border: { style: 'single', size: 4, space: 0, color: { auto: true }, frame: true },
    shading: { pattern: 'clear', fill: { auto: true } },
    vertAlign: 'superscript',
    rtl: true,
    complexScript: true,
    emphasis: 'dot',
    lang: { val: 'tr-TR', eastAsia: 'ja-JP', bidi: 'ar-SA' },
  });
  assert.equal(doc.unknown.size, 0);
});
