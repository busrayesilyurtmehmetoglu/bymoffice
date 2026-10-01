// Regression tests for the findings of the independent review of the docx
// reader: hostile input, field visibility, drawings, and what the model has to
// keep so that a writer can put a document back as it was.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { serializeElement } from '../src/formats/docx/common';
import { readDocx } from '../src/formats/docx/document';
import { NumberingCounters } from '../src/formats/docx/numbering';
import {
  findWrapper,
  inFieldCode,
  paragraphText,
  rawOf,
  safeHyperlinkTarget,
  type Block,
  type DrawingItem,
  type FieldCharItem,
  type LegacyGraphicItem,
  type Paragraph,
  type Run,
  type Table,
} from '../src/model/doc/types';
import { OpcPackage } from '../src/opc';
import { buildDocx, documentXml, readSpec, type DocxSpec } from './docx-fixtures';

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
const texts = (blocks: Block[]) => blocks.filter((block) => block.type === 'paragraph').map((block) => paragraphText(block as Paragraph));

/** Reads the spec and returns the model with the time the reading took. */
async function timed(spec: DocxSpec) {
  const pkg = await OpcPackage.open(await buildDocx(spec));
  const start = performance.now();
  const doc = await readDocx(pkg);
  return { doc, elapsed: performance.now() - start };
}

// ------------------------------------------------------------ hostile input

test('docx review: a million characters of hostile text are read in linear time', async () => {
  const n = 1000000;
  const spaces = ' '.repeat(n);
  const mixed = ' \t\n'.repeat(n / 4);
  // Each of these made a regular expression of the reader (or could make one) back-track.
  const cases: [name: string, spec: DocxSpec, check: (blocks: Block[]) => void][] = [
    [
      'whitespace in the middle of w:t',
      { body: `<w:p><w:r><w:t>x${spaces}x</w:t></w:r></w:p>` },
      (blocks) => assert.equal(paragraphText(para(blocks, 0)).length, n + 2),
    ],
    [
      'whitespace in the middle of w:instrText and w:delText',
      { body: `<w:p><w:r><w:instrText>x${mixed}x</w:instrText><w:delText>x${spaces}x</w:delText></w:r></w:p>` },
      (blocks) => assert.equal((run(para(blocks, 0), 0).content[0] as { text: string }).text.length, mixed.length + 2),
    ],
    [
      'nothing but whitespace, and whitespace on both sides',
      { body: `<w:p><w:r><w:t>${mixed}</w:t><w:t>${spaces}x${spaces}</w:t></w:r></w:p>` },
      (blocks) => assert.equal(paragraphText(para(blocks, 0)), 'x'),
    ],
    [
      'a drawing offset and alignment padded with whitespace',
      {
        body:
          `<w:p><w:r><w:drawing><wp:anchor><wp:positionH relativeFrom="page"><wp:posOffset>1${spaces}x</wp:posOffset></wp:positionH>` +
          `<wp:positionV><wp:align>${spaces}top${spaces}</wp:align><wp:posOffset>${spaces}5${spaces}</wp:posOffset></wp:positionV>` +
          '<wp:extent cx="1" cy="1"/></wp:anchor></w:drawing></w:r></w:p>',
      },
      (blocks) => {
        const drawing = run(para(blocks, 0), 0).content[0] as DrawingItem;
        assert.deepEqual(drawing.anchor?.horizontal, { relativeFrom: 'page' });
        assert.deepEqual(drawing.anchor?.vertical, { offset: 5, align: 'top' });
      },
    ],
    [
      'requirements of an alternate content choice',
      { body: `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps${spaces}w14${spaces}x"><w:t>a</w:t></mc:Choice><mc:Fallback><w:t>b</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>` },
      (blocks) => assert.equal(paragraphText(para(blocks, 0)), 'b'),
    ],
    [
      'measurements and numbers padded with whitespace',
      {
        body:
          `<w:p><w:pPr><w:ind w:left="1${spaces}x" w:right="${spaces}1"/><w:spacing w:before="${'1'.repeat(n)}"/></w:pPr>` +
          `<w:r><w:rPr><w:sz w:val="${'9'.repeat(n)}pt"/><w:w w:val="${'1'.repeat(n)}%"/></w:rPr></w:r></w:p>` +
          `<w:tbl><w:tblPr><w:tblW w:w="${'1'.repeat(n)}.5%" w:type="pct"/></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`,
      },
      (blocks) => assert.deepEqual(para(blocks, 0).props.indent, { right: 1 }),
    ],
  ];
  for (const [name, spec, check] of cases) {
    const { doc, elapsed } = await timed(spec);
    check(doc.body.blocks);
    assert.ok(elapsed < 2000, `${name}: reading took ${elapsed.toFixed(0)} ms`);
  }

  // The number text of a level and the hyperlink check are driven by file content too.
  const percent = '%'.repeat(n) + '%1';
  const start = performance.now();
  const numbered = await readSpec({
    numbering: `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="${percent}"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`,
    body: '<w:p/>',
  });
  assert.equal(new NumberingCounters(numbered.numbering).next(1, 0)!.text.length, n + 1);
  assert.equal(safeHyperlinkTarget('\t'.repeat(n) + 'java\nscript:alert(1)'), undefined);
  assert.equal(safeHyperlinkTarget('https://example.com/' + ' \t'.repeat(n / 2))!.length > 0, true);
  assert.equal(safeHyperlinkTarget('a'.repeat(n) + ':x'), undefined);
  assert.ok(performance.now() - start < 2000);
});

// ------------------------------------------------------------------- fields

const fld = (type: string, extra = '') => `<w:fldChar w:fldCharType="${type}"${extra}/>`;

test('docx review: field code is tracked per item, not per run', async () => {
  const doc = await readSpec({
    body:
      // A: code, separator, result, end and following text share one run.
      `<w:p><w:r>${fld('begin')}</w:r><w:r><w:instrText> PAGE </w:instrText>${fld('separate')}<w:t>3</w:t>${fld('end')}<w:t>after</w:t></w:r></w:p>` +
      // B: a field without separator ends inside a run that goes on with text.
      `<w:p><w:r><w:t>x</w:t></w:r><w:r>${fld('begin')}</w:r><w:r><w:instrText>SEQ a</w:instrText></w:r><w:r>${fld('end')}<w:t>tail</w:t></w:r></w:p>` +
      // C: text before the field in the run that opens it; tabs and breaks in the code are not shown either.
      `<w:p><w:r><w:t>pre</w:t>${fld('begin')}<w:instrText>IF</w:instrText><w:tab/><w:br/><w:t>hidden</w:t></w:r><w:r><w:t>also hidden</w:t>${fld('separate')}<w:tab/><w:t>shown</w:t></w:r><w:r>${fld('end')}</w:r></w:p>`,
  });
  const a = para(doc.body.blocks, 0);
  assert.equal(paragraphText(a), '3after');
  assert.equal(run(a, 0).fieldCode, undefined);
  assert.deepEqual(run(a, 1).fieldCode, [true, true, false, false, false]);
  assert.deepEqual((run(a, 0).content[0] as FieldCharItem).field, { instruction: 'PAGE', result: '3' });
  assert.deepEqual([0, 1, 2, 3, 4].map((i) => inFieldCode(run(a, 1), i)), [true, true, false, false, false]);

  const b = para(doc.body.blocks, 1);
  assert.equal(paragraphText(b), 'xtail');
  assert.equal(run(b, 2).fieldCode, true);
  assert.deepEqual(run(b, 3).fieldCode, [true, false]);
  assert.deepEqual((run(b, 1).content[0] as FieldCharItem).field, { instruction: 'SEQ a', result: '' });

  const c = para(doc.body.blocks, 2);
  assert.equal(paragraphText(c), 'pre\tshown');
  assert.deepEqual(run(c, 0).fieldCode, [false, false, true, true, true, true]);
  assert.deepEqual(run(c, 1).fieldCode, [true, true, false, false]);
  assert.equal(inFieldCode(run(c, 2), 0), false);
  assert.deepEqual((run(c, 0).content[1] as FieldCharItem).field, { instruction: 'IF', result: '\tshown' });
});

test('docx review: a field that is never closed ends with its paragraph', async () => {
  const doc = await readSpec({
    body:
      `<w:p><w:r><w:t>one</w:t></w:r><w:r>${fld('begin')}</w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:t>code</w:t></w:r></w:p>` +
      '<w:p><w:r><w:t>two</w:t></w:r></w:p>' +
      // A field that does close, with its result running on into the next paragraph, still does.
      `<w:p><w:r>${fld('begin')}<w:instrText>TOC</w:instrText>${fld('separate')}<w:t>three</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>four</w:t>${fld('end')}<w:t>five</w:t></w:r></w:p>` +
      // Two fields and one end mark, in a table cell.
      `<w:tbl><w:tr><w:tc><w:p><w:r>${fld('begin')}<w:instrText>A</w:instrText>${fld('begin')}<w:instrText>B</w:instrText></w:r></w:p><w:p><w:r><w:instrText>C</w:instrText>${fld('separate')}<w:t>six</w:t>${fld('end')}</w:r></w:p></w:tc></w:tr></w:tbl>` +
      '<w:p><w:r><w:t>seven</w:t></w:r></w:p>',
    headers: { rIdH: `<w:p><w:r>${fld('begin')}<w:instrText>X</w:instrText></w:r></w:p><w:p><w:r><w:t>header</w:t></w:r></w:p>` },
  });
  assert.deepEqual(texts(doc.body.blocks), ['one', 'two', 'three', 'fourfive', 'seven']);
  const open = run(para(doc.body.blocks, 0), 1).content[0] as FieldCharItem;
  assert.deepEqual(open.field, { instruction: 'PAGE', result: '', unclosed: true });
  const toc = run(para(doc.body.blocks, 2), 0).content[0] as FieldCharItem;
  assert.deepEqual(toc.field, { instruction: 'TOC', result: 'threefour' });

  const cell = (doc.body.blocks[4] as Table).rows[0].cells[0];
  assert.deepEqual(texts(cell.blocks), ['', 'six']);
  const [outer, , inner] = run(para(cell.blocks, 0), 0).content as FieldCharItem[];
  // The one end mark closes the innermost open field, so it is the outer one that never closes.
  assert.deepEqual(outer.field, { instruction: 'A', result: '', unclosed: true });
  assert.deepEqual(inner.field, { instruction: 'BC', result: 'six' });

  assert.deepEqual(texts(doc.headers.get('rIdH')!.blocks), ['', 'header']);

  // More unclosed fields than are assembled: the ones beyond the limit must not wait for an end either.
  const many = await readSpec({
    body:
      `<w:p>${`<w:r>${fld('begin')}</w:r>`.repeat(40)}<w:r><w:t>code</w:t></w:r></w:p>` +
      `<w:p><w:r>${fld('begin')}<w:instrText>PAGE</w:instrText>${fld('separate')}<w:t>7</w:t>${fld('end')}<w:t>after</w:t></w:r></w:p>`,
  });
  assert.deepEqual(texts(many.body.blocks), ['', '7after']);
  const begins = para(many.body.blocks, 0).children.slice(0, 40).map((inline) => (inline as Run).content[0] as FieldCharItem);
  assert.equal(begins.filter((item) => item.field?.unclosed).length, 32);
  assert.equal(begins.filter((item) => item.field === undefined).length, 8);
  assert.deepEqual((run(para(many.body.blocks, 1), 0).content[0] as FieldCharItem).field, { instruction: 'PAGE', result: '7' });
});

test('docx review: a field character with attributes the reader does not know still drives the field', async () => {
  const doc = await readSpec({
    body:
      `<w:p><w:r>${fld('begin', ' w14:x="1"')}</w:r><w:r><w:instrText>DATE</w:instrText></w:r><w:r>${fld('separate', ' w:dirty="maybe"')}</w:r>` +
      `<w:r><w:t>today</w:t></w:r><w:r>${fld('end', ' w:foo="bar"')}</w:r><w:r><w:t>rest</w:t></w:r></w:p>` +
      '<w:p><w:r><w:t>next</w:t></w:r></w:p>' +
      // A type the reader cannot interpret is kept as it is and is no field character.
      '<w:p><w:r><w:fldChar w:fldCharType="middle"/><w:fldChar/><w:t>plain</w:t></w:r></w:p>',
  });
  const p = para(doc.body.blocks, 0);
  assert.deepEqual(texts(doc.body.blocks), ['todayrest', 'next', 'plain']);
  const begin = run(p, 0).content[0] as FieldCharItem;
  assert.equal(begin.type, 'fieldChar');
  assert.deepEqual(begin.field, { instruction: 'DATE', result: 'today' });
  assert.equal(serializeElement(begin.xml!), '<w:fldChar w:fldCharType="begin" w14:x="1"/>');
  const separate = run(p, 2).content[0] as FieldCharItem;
  assert.equal(separate.charType, 'separate');
  assert.equal(separate.dirty, undefined);
  assert.equal(separate.xml?.attrs.get('w:dirty'), 'maybe');
  assert.equal((run(p, 4).content[0] as FieldCharItem).xml?.attrs.get('w:foo'), 'bar');
  // A field character the model holds in full carries no copy of the element.
  const plain = await readSpec({ body: `<w:p><w:r>${fld('begin', ' w:dirty="1" w:fldLock="0"')}</w:r><w:r>${fld('end')}</w:r></w:p>` });
  assert.deepEqual(run(para(plain.body.blocks, 0), 0).content[0], { type: 'fieldChar', charType: 'begin', dirty: true, locked: false, field: { instruction: '', result: '' } });
  assert.deepEqual(types(run(para(doc.body.blocks, 2), 0).content), ['unknown', 'unknown', 'text']);
  assert.equal(doc.unknown.get('r/w:fldChar'), 2);
});

// ----------------------------------------------------------------- drawings

const WPG = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const box = (text: string) => `<w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent>`;
const pic = (rId: string, extra = '', xfrm = '') =>
  `<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="${rId}"/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}"/>${extra}</pic:blipFill><pic:spPr>${xfrm}</pic:spPr></pic:pic>`;

test('docx review: alternate content inside a drawing is read through the selected branch only', async () => {
  const group = (requires: string) =>
    `<w:drawing><wp:inline><wp:extent cx="100" cy="100"/><a:graphic><a:graphicData uri="${WPG}"><wpg:wgp xmlns:wpg="${WPG}">` +
    `<mc:AlternateContent><mc:Choice Requires="${requires}"><wps:wsp><wps:txbx>${box('A')}</wps:txbx></wps:wsp></mc:Choice>` +
    `<mc:Fallback><w:pict><v:textbox>${box('A (fallback)')}</v:textbox><v:imagedata r:id="rImg2"/></w:pict></mc:Fallback></mc:AlternateContent>` +
    `<wps:wsp><wps:txbx>${box('B')}</wps:txbx></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>`;
  const doc = await readSpec({
    rels: [['rImg2', 'image', 'media/image2.png']],
    body:
      `<w:p><w:r>${group('wps')}</w:r><w:r>${group('unknownThing')}</w:r>` +
      // The same inside a VML picture.
      `<w:r><w:pict><v:group><mc:AlternateContent><mc:Choice Requires="wps"><v:textbox>${box('C')}</v:textbox></mc:Choice>` +
      `<mc:Fallback><v:textbox>${box('C (fallback)')}</v:textbox></mc:Fallback></mc:AlternateContent></v:group></w:pict></w:r></w:p>`,
  });
  const p = para(doc.body.blocks, 0);
  const boxes = (index: number) => (run(p, index).content[0] as DrawingItem | LegacyGraphicItem).textBoxes!.map((blocks) => texts(blocks).join());
  assert.deepEqual(boxes(0), ['A', 'B']);
  assert.deepEqual(boxes(1), ['A (fallback)', 'B']);
  assert.deepEqual(boxes(2), ['C']);
});

test('docx review: every picture of a drawing is in the model, each with its own cropping and place', async () => {
  const xfrm = (x: number, y: number, cx: number, cy: number, attrs = '') => `<a:xfrm${attrs}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;
  const doc = await readSpec({
    rels: [
      ['rImg1', 'image', 'media/image1.png'],
      ['rImg2', 'image', 'media/image2.png'],
      ['rImg3', 'image', 'media/image3.png'],
      ['rImg4', 'image', 'https://example.com/4.png', true],
    ],
    body:
      `<w:p><w:r><w:drawing><wp:inline><wp:extent cx="2000" cy="1000"/><a:graphic><a:graphicData uri="${WPG}"><wpg:wgp xmlns:wpg="${WPG}">` +
      `<wpg:grpSpPr><a:xfrm rot="5400000" flipH="1"><a:off x="0" y="0"/><a:ext cx="2000" cy="1000"/><a:chOff x="10" y="20"/><a:chExt cx="4000" cy="2000"/></a:xfrm></wpg:grpSpPr>` +
      // The first picture has no cropping; the second has. The cropping must not move to the first.
      pic('rImg1', '', xfrm(10, 20, 1000, 2000)) +
      pic('rImg2', '<a:srcRect l="1000" b="2500"/>', xfrm(1010, 20, 3000, 2000, ' rot="-60000" flipV="true"')) +
      `<wpg:grpSp><wpg:grpSpPr>${xfrm(0, 0, 5, 5)}</wpg:grpSpPr>${pic('rImg3')}` +
      // A picture used as the fill of a shape, with a linked image.
      `<wps:wsp xmlns:wps="${WPS}"><wps:spPr>${xfrm(1, 2, 3, 4)}<a:blipFill><a:blip r:link="rImg4"/><a:srcRect t="5"/></a:blipFill></wps:spPr>` +
      // A picture inside a text box belongs to the drawing in that text box.
      `<wps:txbx><w:txbxContent><w:p><w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><a:graphic><a:graphicData uri="${PIC}">${pic('rImg1')}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:txbxContent></wps:txbx>` +
      '</wps:wsp></wpg:grpSp>' +
      // No relationship id: not a picture the model can show.
      '<pic:pic><pic:blipFill><a:blip/></pic:blipFill></pic:pic>' +
      '</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>',
  });
  const drawing = run(para(doc.body.blocks, 0), 0).content[0] as DrawingItem;
  assert.equal(drawing.kind, 'group');
  const group = { x: 0, y: 0, width: 2000, height: 1000, rotation: 5400000, flipH: true, childX: 10, childY: 20, childWidth: 4000, childHeight: 2000 };
  assert.deepEqual(drawing.transform, group);
  assert.deepEqual(drawing.images, [
    { image: { rId: 'rImg1', target: '/word/media/image1.png' }, name: 'rImg1', transform: { x: 10, y: 20, width: 1000, height: 2000 }, groups: [group] },
    {
      image: { rId: 'rImg2', target: '/word/media/image2.png' },
      name: 'rImg2',
      crop: { left: 1000, bottom: 2500 },
      transform: { x: 1010, y: 20, width: 3000, height: 2000, rotation: -60000, flipV: true },
      groups: [group],
    },
    { image: { rId: 'rImg3', target: '/word/media/image3.png' }, name: 'rImg3', groups: [group, { x: 0, y: 0, width: 5, height: 5 }] },
    {
      image: { rId: 'rImg4', linked: true, target: 'https://example.com/4.png', external: true },
      crop: { top: 5 },
      fill: true,
      transform: { x: 1, y: 2, width: 3, height: 4 },
      groups: [group, { x: 0, y: 0, width: 5, height: 5 }],
    },
  ]);
  // The single-image fields describe the first picture.
  assert.deepEqual(drawing.image, { rId: 'rImg1', target: '/word/media/image1.png' });
  assert.equal(drawing.crop, undefined);
  const inner = run(drawing.textBoxes![0][0] as Paragraph, 0).content[0] as DrawingItem;
  assert.equal(inner.images?.length, 1);
});

test('docx review: what layout needs from a drawing is read', async () => {
  const doc = await readSpec({
    rels: [['rImg1', 'image', 'media/image1.png']],
    body:
      '<w:p><w:r><w:drawing><wp:inline distT="1" distB="2" distL="3" distR="4"><wp:extent cx="100" cy="200"/><wp:effectExtent l="5" t="6" r="7" b="8"/>' +
      `<a:graphic><a:graphicData uri="${PIC}">${pic('rImg1', '', '<a:xfrm rot="2700000" flipH="1" flipV="0"><a:off x="0" y="0"/><a:ext cx="100" cy="200"/></a:xfrm>')}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>` +
      '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="10" distR="10" simplePos="0" relativeHeight="5" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><mc:AlternateContent><mc:Choice Requires="wp14"><wp14:pctPosHOffset>25000</wp14:pctPosHOffset></mc:Choice>' +
      '<mc:Fallback><wp:posOffset>123</wp:posOffset></mc:Fallback></mc:AlternateContent></wp:positionH>' +
      '<wp:positionV relativeFrom="page"><wp14:pctPosVOffset>-5000</wp14:pctPosVOffset></wp:positionV><wp:extent cx="300" cy="400"/><wp:effectExtent l="0" t="0" r="9" b="0"/>' +
      '<wp:wrapTight wrapText="left" distL="11" distR="12"><wp:wrapPolygon edited="1"><wp:start x="0" y="0"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="21600" y="-21600"/><wp:lineTo x="bad" y="0"/></wp:wrapPolygon></wp:wrapTight>' +
      `<wp:docPr id="2" name="s"/><a:graphic><a:graphicData uri="${WPS}"><wps:wsp><wps:spPr><a:xfrm flipV="1"><a:off x="0" y="0"/><a:ext cx="300" cy="400"/></a:xfrm></wps:spPr></wps:wsp></a:graphicData></a:graphic>` +
      '<wp14:sizeRelH relativeFrom="page"><wp14:pctWidth>50000</wp14:pctWidth></wp14:sizeRelH><wp14:sizeRelV relativeFrom="margin"><wp14:pctHeight>0</wp14:pctHeight></wp14:sizeRelV>' +
      '</wp:anchor></w:drawing></w:r></w:p>',
  });
  const p = para(doc.body.blocks, 0);
  const inline = run(p, 0).content[0] as DrawingItem;
  assert.deepEqual(inline.inline, { distT: 1, distB: 2, distL: 3, distR: 4 });
  assert.deepEqual(inline.effectExtent, { left: 5, top: 6, right: 7, bottom: 8 });
  assert.deepEqual(inline.transform, { x: 0, y: 0, width: 100, height: 200, rotation: 2700000, flipH: true, flipV: false });
  assert.equal(inline.anchor, undefined);

  const anchored = run(p, 1).content[0] as DrawingItem;
  assert.equal(anchored.inline, undefined);
  assert.deepEqual(anchored.effectExtent, { left: 0, top: 0, right: 9, bottom: 0 });
  assert.deepEqual(anchored.transform, { x: 0, y: 0, width: 300, height: 400, flipV: true });
  assert.deepEqual(anchored.anchor, {
    distT: 0,
    distB: 0,
    distL: 10,
    distR: 10,
    relativeHeight: 5,
    behindDoc: false,
    locked: false,
    layoutInCell: true,
    allowOverlap: true,
    // The choice the reader understands is taken: a position in thousandths of a percent.
    horizontal: { relativeFrom: 'margin', pctOffset: 25000 },
    vertical: { relativeFrom: 'page', pctOffset: -5000 },
    wrap: 'tight',
    wrapText: 'left',
    wrapDist: { left: 11, right: 12 },
    wrapPolygon: {
      edited: true,
      points: [
        [0, 0],
        [21600, 0],
        [21600, -21600],
      ],
    },
    sizeRelH: { relativeFrom: 'page', pct: 50000 },
    sizeRelV: { relativeFrom: 'margin', pct: 0 },
  });
});

// ------------------------------------------------------------- lossless gaps

test('docx review: both spellings of an edge are kept apart', async () => {
  const doc = await readSpec({
    body:
      '<w:p><w:pPr><w:ind w:left="100" w:start="200"/></w:pPr></w:p>' +
      '<w:p><w:pPr><w:ind w:start="200" w:end="300" w:hanging="5"/><w:jc w:val="start"/><w:pBdr><w:start w:val="single"/><w:right w:val="double"/></w:pBdr></w:pPr></w:p>' +
      '<w:p><w:pPr><w:ind w:left="200" w:right="300"/><w:jc w:val="left"/></w:pPr></w:p>' +
      '<w:p><w:pPr><w:ind w:left="200" w:end="300"/></w:pPr></w:p>' +
      '<w:p><w:pPr><w:ind w:startChars="10" w:endChars="20"/></w:pPr></w:p>' +
      '<w:p><w:pPr><w:pBdr><w:left w:val="single"/><w:start w:val="double"/></w:pBdr></w:pPr></w:p>' +
      '<w:tbl><w:tblPr><w:jc w:val="end"/><w:tblBorders><w:start w:val="single"/><w:end w:val="dotted"/></w:tblBorders>' +
      '<w:tblCellMar><w:start w:w="10" w:type="dxa"/><w:right w:w="20" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
      '<w:tr><w:tc><w:tcPr><w:tcBorders><w:end w:val="single"/></w:tcBorders><w:tcMar><w:end w:w="30" w:type="dxa"/></w:tcMar></w:tcPr><w:p/></w:tc></w:tr></w:tbl>',
  });
  const props = (i: number) => para(doc.body.blocks, i).props;
  // Two values for one edge: the first is the one the model shows, the element is kept for the other.
  assert.deepEqual(props(0).indent, { left: 100 });
  assert.deepEqual(props(0).extras?.map((extra) => [extra.index, extra.el.name, extra.partial]), [[0, 'w:ind', true]]);
  assert.equal(props(0).extras![0].el.attrs.get('w:start'), '200');

  assert.deepEqual(props(1).indent, { left: 200, right: 300, hanging: 5, logical: true });
  assert.equal(props(1).align, 'start');
  assert.deepEqual(props(1).borders, { left: { style: 'single', logical: true }, right: { style: 'double' } });
  assert.equal(props(1).extras, undefined);
  assert.deepEqual(props(2).indent, { left: 200, right: 300 });
  assert.equal(props(2).align, 'left');
  // One edge in each spelling: no single flag says that, so the element is kept.
  assert.deepEqual(props(3).indent, { left: 200, right: 300 });
  assert.deepEqual(props(3).extras?.map((extra) => extra.partial), [true]);
  assert.deepEqual(props(4).indent, { leftChars: 10, rightChars: 20, logical: true });
  assert.deepEqual(props(5).borders, { left: { style: 'single' } });
  assert.deepEqual(props(5).extras?.map((extra) => [extra.el.name, extra.partial]), [['w:pBdr', true]]);

  const table = doc.body.blocks[6] as Table;
  assert.equal(table.props.align, 'end');
  assert.deepEqual(table.props.borders, { left: { style: 'single', logical: true }, right: { style: 'dotted', logical: true } });
  assert.deepEqual(table.props.cellMargins, { left: { type: 'dxa', value: 10, logical: true }, right: { type: 'dxa', value: 20 } });
  assert.equal(table.props.extras, undefined);
  const cell = table.rows[0].cells[0].props;
  assert.deepEqual(cell.borders, { right: { style: 'single', logical: true } });
  assert.deepEqual(cell.margins, { right: { type: 'dxa', value: 30, logical: true } });
  assert.equal(doc.unknown.size, 2);
});

test('docx review: values the reader used to replace or invent are kept as written', async () => {
  const doc = await readSpec({
    body:
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="84A1"/></w:tblPr>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="1234" w:type="auto"/><w:shd w:fill="FF0000"/></w:tcPr><w:p/></w:tc>' +
      '<w:tc><w:tcPr><w:tcW w:type="nil"/><w:shd w:val="clear" w:color="auto"/></w:tcPr><w:p/></w:tc>' +
      '<w:tc><w:tcPr><w:tcW w:w="abc" w:type="auto"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>' +
      '<w:tbl><w:tblPr><w:tblLook w:firstRow="1" w:lastColumn="true" w:noVBand="0"/></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>' +
      '<w:tbl><w:tblPr><w:tblLook w:val="01E0" w:firstRow="0" w:noHBand="1"/></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>' +
      '<w:tbl><w:tblPr><w:tblLook w:val="0600"/></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>' +
      '<w:tbl><w:tblPr><w:tblLook w:val="xyz"/></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>' +
      '<w:p><w:pPr><w:shd w:themeFill="accent1"/></w:pPr><w:r><w:rPr><w:shd w:val="pct10"/></w:rPr></w:r></w:p>',
  });
  const table = (i: number) => doc.body.blocks[i] as Table;
  const cells = table(0).rows[0].cells;
  assert.deepEqual(table(0).props.width, { type: 'auto', value: 0 });
  assert.deepEqual(cells[0].props.width, { type: 'auto', value: 1234 });
  assert.deepEqual(cells[1].props.width, { type: 'nil', value: 0 });
  // A width that is not a number is not understood; the type alone is read and the element kept.
  assert.deepEqual(cells[2].props.width, { type: 'auto', value: 0 });
  assert.deepEqual(cells[2].props.extras?.map((extra) => [extra.el.name, extra.partial]), [['w:tcW', true]]);
  // No pattern was written, so none is stated.
  assert.deepEqual(cells[0].props.shading, { fill: { rgb: 'FF0000' } });
  assert.deepEqual(cells[1].props.shading, { pattern: 'clear', color: { auto: true } });
  const p = para(doc.body.blocks, 5);
  assert.deepEqual(p.props.shading, { fill: { theme: 'accent1' } });
  assert.deepEqual(run(p, 0).props.shading, { pattern: 'pct10' });

  const off = { firstRow: false, lastRow: false, firstColumn: false, lastColumn: false, noHBand: false, noVBand: false };
  // The bits outside the six flags (0x8001 here) and the spelling used are kept.
  assert.deepEqual(table(0).props.look, { ...off, firstRow: true, firstColumn: true, noVBand: true, val: 0x84a1 });
  assert.deepEqual(table(1).props.look, { ...off, firstRow: true, lastColumn: true, stated: 0x0020 | 0x0100 | 0x0400 });
  // Each flag of the bit field, one by one; the separate attributes win over it.
  assert.deepEqual(table(2).props.look, { ...off, lastRow: true, firstColumn: true, lastColumn: true, noHBand: true, val: 0x01e0, stated: 0x0020 | 0x0200 });
  assert.deepEqual(table(3).props.look, { ...off, noHBand: true, noVBand: true, val: 0x0600 });
  assert.equal(table(4).props.look, undefined);
  assert.deepEqual(table(4).props.extras?.map((extra) => extra.el.name), ['w:tblLook']);
});

test('docx review: content of cell-level and row-level wrappers keeps its order and its wrapper', async () => {
  const cell = (text: string) => `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const doc = await readSpec({
    body:
      '<w:tbl><w:tr>' +
      cell('a') +
      `<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent>${cell('b')}<w:bookmarkEnd w:id="7"/>${cell('c')}</w:sdtContent></w:sdt>` +
      '<w:bookmarkStart w:id="8" w:name="x"/>' +
      `<w:customXml w:element="e">${cell('d')}</w:customXml>` +
      '</w:tr>' +
      '<w:sdt><w:sdtContent><w:bookmarkStart w:id="9" w:name="y"/><w:tr>' +
      cell('e') +
      '</w:tr><w:bookmarkEnd w:id="9"/></w:sdtContent></w:sdt>' +
      // A wrapper that holds no row at all is not lost either.
      '<w:sdt><w:sdtPr><w:id w:val="2"/></w:sdtPr><w:sdtContent><w:bookmarkEnd w:id="8"/></w:sdtContent></w:sdt>' +
      '</w:tbl>',
  });
  const table = doc.body.blocks[0] as Table;
  const [first, second] = table.rows;
  assert.equal(first.cells.length, 4);
  // The wrapper of a cell-level content control reaches each of its cells.
  const sdt = first.cells[1].wrap;
  assert.equal(sdt?.kind, 'sdt');
  assert.equal(first.cells[0].wrap, undefined);
  assert.equal(first.cells[2].wrap, sdt);
  assert.equal(first.cells[3].wrap?.kind, 'customXml');
  assert.deepEqual(
    first.extras?.map((extra) => [extra.index, extra.el.name, extra.wrap === sdt, extra.wrap === undefined]),
    [
      // Second of the three children of the content control.
      [1, 'w:bookmarkEnd', true, false],
      // Third child of the row.
      [2, 'w:bookmarkStart', false, true],
    ],
  );

  const rowWrap = second.wrap;
  assert.equal(rowWrap?.kind, 'sdt');
  assert.deepEqual(
    table.extras?.map((extra) => [extra.index, extra.el.name, extra.wrap === rowWrap]),
    [
      [0, 'w:bookmarkStart', true],
      [2, 'w:bookmarkEnd', true],
      [0, 'w:bookmarkEnd', false],
    ],
  );
  const wrap = table.extras![2].wrap;
  const lone = wrap && wrap.kind === 'sdt' ? wrap : undefined;
  assert.ok(lone !== undefined && lone !== rowWrap);
  assert.equal(serializeElement(lone!.pr!), '<w:sdtPr><w:id w:val="2"/></w:sdtPr>');
});

test('docx review: comments and processing instructions stay where they were', async () => {
  const body =
    '<!--before--><w:p><!--in p--><w:r><?run pi?><w:t>a</w:t><!--in r--></w:r><w:hyperlink w:anchor="x"><!--in link--><w:r><w:t>b</w:t></w:r></w:hyperlink></w:p>' +
    '<?between a="1"?><w:tbl><!--in tbl--><w:tr><!--in tr--><w:tc><!--in tc--><w:p/><w:sdt><w:sdtPr/><!--in sdt--><w:sdtContent><!--in content--><w:p/></w:sdtContent></w:sdt></w:tc>' +
    '<w:sdt><w:sdtContent><w:tc><w:p/></w:tc><!--in cell sdt--></w:sdtContent></w:sdt></w:tr><!--after tr--></w:tbl>' +
    '<w:p><w:pPr><w:jc w:val="center"/><!--in pPr--></w:pPr><w:r><w:rPr><w:b><!--in b--></w:b></w:rPr><w:t>c</w:t></w:r></w:p>' +
    '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>c</w:t></w:r></w:p>' +
    '<w:sectPr><!--in sectPr--><w:pgSz w:w="100" w:h="200"/></w:sectPr><!--last-->';
  const document =
    '<?xml version="1.0" encoding="UTF-8"?><!--prolog--><?app x?>' +
    documentXml(body).replace(/^<\?xml[^>]*>/, '').replace('<w:body>', '<!--in document--><w:body>') +
    '<!--epilog-->';
  const doc = await readSpec({
    document,
    footnotes: '<!--first--><w:footnote w:id="1"><!--in note--><w:p/></w:footnote><?last?>',
  });
  const blocks = doc.body.blocks;
  assert.deepEqual(types(blocks), ['misc', 'paragraph', 'misc', 'table', 'paragraph', 'paragraph', 'misc']);
  assert.deepEqual(blocks[0], { type: 'misc', node: { type: 'comment', value: 'before' }, srcStart: blocks[0].srcStart, srcEnd: blocks[0].srcEnd });
  assert.equal(rawOf(doc.body, blocks[0]), '<!--before-->');
  assert.deepEqual((blocks[2] as { node: unknown }).node, { type: 'pi', target: 'between', value: 'a="1"' });
  assert.equal(rawOf(doc.body, blocks[2]), '<?between a="1"?>');
  assert.equal(rawOf(doc.body, blocks[6]), '<!--last-->');
  assert.equal(doc.body.prolog, '<?xml version="1.0" encoding="UTF-8"?><!--prolog--><?app x?>');
  assert.equal(doc.body.epilog, '<!--epilog-->');
  assert.deepEqual(doc.body.misc, [{ index: 0, node: { type: 'comment', value: 'in document' } }]);

  const p = para(blocks, 1);
  assert.deepEqual(types(p.children), ['misc', 'run', 'misc', 'run']);
  assert.deepEqual(types(run(p, 1).content), ['misc', 'text', 'misc']);
  assert.equal(p.children[2].wrap?.kind, 'hyperlink');
  assert.equal(p.children[2].wrap, p.children[3].wrap);
  assert.equal(paragraphText(p), 'ab');

  const table = blocks[3] as Table;
  assert.deepEqual(table.misc, [
    { index: 0, node: { type: 'comment', value: 'in tbl' } },
    { index: 1, node: { type: 'comment', value: 'after tr' } },
  ]);
  const row = table.rows[0];
  assert.deepEqual(row.misc?.map((item) => [item.index, item.node.type === 'comment' && item.node.value, item.wrap?.kind]), [
    [0, 'in tr', undefined],
    [1, 'in cell sdt', 'sdt'],
  ]);
  const cell = row.cells[0];
  assert.deepEqual(types(cell.blocks), ['misc', 'paragraph', 'misc', 'paragraph']);
  const wrapper = findWrapper(cell.blocks[3], 'sdt')!;
  assert.equal(cell.blocks[2].wrap, wrapper);
  assert.deepEqual(wrapper.misc, [{ index: 1, node: { type: 'comment', value: 'in sdt' } }]);

  // Properties with a comment in them are kept whole and are not shared with the same properties without one.
  const [withComment, without] = [para(blocks, 4), para(blocks, 5)];
  assert.notEqual(withComment.props, without.props);
  assert.deepEqual(withComment.props.extras?.map((extra) => [extra.index, extra.partial]), [[-1, true]]);
  assert.equal(without.props.extras, undefined);
  assert.notEqual(run(withComment, 0).props, run(without, 0).props);
  assert.deepEqual(run(withComment, 0).props.extras?.map((extra) => [extra.index, extra.el.name, extra.partial]), [[0, 'w:b', true]]);
  assert.equal(run(withComment, 0).props.bold, true);
  assert.equal(serializeElement(withComment.props.extras![0].el), '<w:pPr><w:jc w:val="center"/><!--in pPr--></w:pPr>');
  assert.deepEqual(doc.body.section?.extras?.map((extra) => [extra.index, extra.partial]), [[-1, true]]);

  const notes = doc.footnotes!;
  assert.deepEqual(notes.misc, [
    { index: 0, node: { type: 'comment', value: 'first' } },
    { index: 1, node: { type: 'pi', target: 'last', value: '' } },
  ]);
  assert.deepEqual(types(notes.items.get(1)!.blocks), ['misc', 'paragraph']);
});

test('docx review: the place of a section mark that is not last, and odd children of a notes part, are kept', async () => {
  const doc = await readSpec({
    body: '<w:p/><w:sectPr><w:pgSz w:w="100" w:h="200"/></w:sectPr><w:p><w:r><w:t>after</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="300" w:h="400"/></w:sectPr>',
    footnotes:
      '<w:footnote w:id="2"><w:p><w:r><w:t>first</w:t></w:r></w:p></w:footnote>' +
      '<w:footnote w:id="2"><w:p><w:r><w:t>again</w:t></w:r></w:p></w:footnote>' +
      '<w:other w:x="1"><w:p/></w:other>' +
      '<w:footnote><w:p/></w:footnote>' +
      '<w:footnote w:id="x"/>' +
      '<w:footnote w:id="3"><w:p/></w:footnote>',
    comments: '<w:comment w:id="1"/><w:comment w:id="1" w:author="B"/>',
  });
  // The first section mark is the one of the body; it sat before the second block.
  assert.deepEqual(doc.body.section?.page, { width: 100, height: 200 });
  assert.equal(doc.body.sectionIndex, 1);
  assert.equal(rawOf(doc.body, doc.body.sectionSpan), '<w:sectPr><w:pgSz w:w="100" w:h="200"/></w:sectPr>');
  assert.deepEqual(types(doc.body.blocks), ['paragraph', 'paragraph', 'unknown']);
  const usual = await readSpec({ body: '<w:p/><w:sectPr/>' });
  assert.equal(usual.body.sectionIndex, undefined);
  assert.ok(usual.body.section);

  const notes = doc.footnotes!;
  assert.deepEqual([...notes.items.keys()], [2, 3]);
  assert.equal(paragraphText(notes.items.get(2)!.blocks[0] as Paragraph), 'first');
  assert.deepEqual(
    notes.extras?.map((extra) => [extra.index, serializeElement(extra.el)]),
    [
      [1, '<w:footnote w:id="2"><w:p><w:r><w:t>again</w:t></w:r></w:p></w:footnote>'],
      [2, '<w:other w:x="1"><w:p/></w:other>'],
      [3, '<w:footnote><w:p/></w:footnote>'],
      [4, '<w:footnote w:id="x"/>'],
    ],
  );
  assert.equal(rawOf(notes, notes.items.get(3)!), '<w:footnote w:id="3"><w:p/></w:footnote>');
  assert.deepEqual(doc.comments!.extras?.map((extra) => extra.index), [1]);
  // Each of the three notes and the comment that could not get a place by id is reported.
  assert.equal(doc.warnings.length, 4);
  assert.equal(doc.unknown.get('footnotes/w:other'), 1);
});

test('docx review: a repeated table style condition is kept as not understood', async () => {
  const doc = await readSpec({
    body: '<w:p/>',
    styles:
      '<w:style w:type="table" w:styleId="T">' +
      '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr>' +
      '<w:tblStylePr w:type="firstRow"><w:rPr><w:i/></w:rPr></w:tblStylePr>' +
      '<w:tblStylePr w:type="lastRow"><w:rPr><w:i/></w:rPr><w:mystery/></w:tblStylePr>' +
      '<w:tblStylePr><w:rPr><w:i/></w:rPr></w:tblStylePr>' +
      '<w:tblStylePr w:type="firstCol" w:x="1"><w:rPr><w:i/></w:rPr></w:tblStylePr>' +
      '</w:style>',
  });
  const style = doc.styles.styles.get('T')!;
  assert.deepEqual([...style.conditions!.keys()], ['firstRow', 'lastRow', 'firstCol']);
  assert.deepEqual(style.conditions!.get('firstRow'), { run: { bold: true } });
  assert.deepEqual(
    style.extras?.map((extra) => [extra.index, extra.partial === true]),
    [
      // The repeated one is not in the model at all.
      [1, false],
      // In the model, with a child the model does not hold.
      [2, true],
      [3, false],
      // In the model, with an attribute the model does not hold.
      [4, true],
    ],
  );
});

test('docx review: text that was trimmed keeps what was written', async () => {
  const doc = await readSpec({
    body:
      '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>  padded\n</w:t><w:t>tail </w:t><w:t>\thead</w:t><w:t>plain</w:t><w:t xml:space="preserve"> kept </w:t><w:instrText> PAGE </w:instrText><w:delText>\tgone</w:delText></w:r></w:p></w:tc></w:tr></w:tbl>',
  });
  const content = run((doc.body.blocks[0] as Table).rows[0].cells[0].blocks[0] as Paragraph, 0).content;
  assert.deepEqual(content, [
    { type: 'text', text: 'padded', source: '  padded\n' },
    { type: 'text', text: 'tail', source: 'tail ' },
    { type: 'text', text: 'head', source: '\thead' },
    { type: 'text', text: 'plain' },
    { type: 'text', text: ' kept ' },
    { type: 'instrText', text: 'PAGE', source: ' PAGE ' },
    { type: 'deletedText', text: 'gone', source: '\tgone' },
  ]);
});

test('docx review: source text is kept once and blocks point into it', async () => {
  const first = '<w:p><w:r><w:t>one</w:t></w:r></w:p>';
  const sdt = '<w:sdt><w:sdtContent><w:p/></w:sdtContent></w:sdt>';
  const doc = await readSpec({
    // Line endings are normalised when the part is read; the offsets are into the normalised text.
    // Nothing on the way may throw the offsets off: a ">" in an attribute value of either
    // quoting, or markup inside a CDATA section (which is text, and text between blocks is not content).
    document:
      '\uFEFF' +
      documentXml(`${first}\r\n  <![CDATA[<w:p>not a block</w:p>]]>${sdt}<w:p\r\n w:rsidR="1"/>`)
        .replace('<w:document ', `<w:document w:a='>' w:b=">" `)
        .replace('<w:body>', `<w:body w:c='"/>' w:d="'>">`),
    headers: { rIdH: '<w:p><w:r><w:t>header</w:t></w:r></w:p>' },
    footnotes: `<w:footnote w:id='5' w:x='a>b"'/><w:footnote w:id="6" w:x="a>b'"><w:p/></w:footnote>`,
  });
  // A note is one element in the source, also when it is empty.
  const notes = doc.footnotes!;
  assert.equal(rawOf(notes, notes.items.get(5)), `<w:footnote w:id='5' w:x='a>b"'/>`);
  assert.equal(rawOf(notes, notes.items.get(6)), `<w:footnote w:id="6" w:x="a>b'"><w:p/></w:footnote>`);
  const [one, wrapped, last] = doc.body.blocks;
  assert.equal(rawOf(doc.body, one), first);
  assert.equal(rawOf(doc.body, wrapped), undefined);
  assert.equal(rawOf(doc.body, wrapped.wrap), sdt);
  assert.equal(rawOf(doc.body, last), '<w:p\n w:rsidR="1"/>');
  assert.equal(doc.body.source.indexOf(first), one.srcStart);
  assert.ok(!doc.body.source.includes('\r') && doc.body.source.charCodeAt(0) !== 0xfeff);
  const header = doc.headers.get('rIdH')!;
  assert.equal(rawOf(header, header.blocks[0]), '<w:p><w:r><w:t>header</w:t></w:r></w:p>');
  assert.equal(rawOf(doc.body, undefined), undefined);
  assert.equal(rawOf(doc.body, {}), undefined);
});

// --------------------------------------------------------- layout properties

test('docx review: frames, floating tables and section details are in the model', async () => {
  const doc = await readSpec({
    body:
      '<w:p><w:pPr><w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="page" w:hSpace="57"/></w:pPr></w:p>' +
      '<w:p><w:pPr><w:framePr w:w="2000" w:h="1000" w:hRule="exact" w:vSpace="10" w:x="-5" w:y="7" w:xAlign="center" w:yAlign="top" w:anchorLock="1" w:bogus="1"/></w:pPr></w:p>' +
      '<w:tbl><w:tblPr><w:tblpPr w:leftFromText="141" w:rightFromText="142" w:topFromText="1" w:bottomFromText="2" w:vertAnchor="page" w:horzAnchor="margin" w:tblpX="-10" w:tblpY="20" w:tblpXSpec="center" w:tblpYSpec="bottom"/>' +
      '<w:tblOverlap w:val="never"/></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>' +
      '<w:sectPr><w:footnotePr><w:pos w:val="beneathText"/><w:numFmt w:val="lowerRoman"/><w:numStart w:val="3"/><w:numRestart w:val="eachPage"/></w:footnotePr>' +
      '<w:endnotePr><w:numFmt w:val="decimal"/><w:mystery/></w:endnotePr>' +
      '<w:pgBorders w:offsetFrom="page" w:display="firstPage" w:zOrder="back"><w:top w:val="single" w:sz="4" w:space="24" w:color="auto"/><w:left w:val="double" w:sz="8"/><w:bottom w:val="apples" w:sz="12" r:id="rId9"/></w:pgBorders>' +
      '<w:lnNumType w:countBy="5" w:start="2" w:distance="360" w:restart="newSection"/></w:sectPr>',
  });
  assert.deepEqual(para(doc.body.blocks, 0).props, { frame: { dropCap: 'drop', lines: 3, wrap: 'around', vAnchor: 'text', hAnchor: 'page', hSpace: 57 } });
  const framed = para(doc.body.blocks, 1).props;
  assert.deepEqual(framed.frame, { width: 2000, height: 1000, heightRule: 'exact', vSpace: 10, x: -5, y: 7, xAlign: 'center', yAlign: 'top', anchorLock: true });
  assert.deepEqual(framed.extras?.map((extra) => [extra.el.name, extra.partial]), [['w:framePr', true]]);

  const table = doc.body.blocks[2] as Table;
  assert.deepEqual(table.props, {
    floating: {
      leftFromText: 141,
      rightFromText: 142,
      topFromText: 1,
      bottomFromText: 2,
      vertAnchor: 'page',
      horzAnchor: 'margin',
      x: -10,
      y: 20,
      xAlign: 'center',
      yAlign: 'bottom',
    },
    overlap: 'never',
  });

  const section = doc.body.section!;
  assert.deepEqual(section.footnotes, { position: 'beneathText', format: 'lowerRoman', start: 3, restart: 'eachPage' });
  assert.deepEqual(section.endnotes, { format: 'decimal' });
  assert.deepEqual(section.pageBorders, {
    offsetFrom: 'page',
    display: 'firstPage',
    zOrder: 'back',
    top: { style: 'single', size: 4, space: 24, color: { auto: true } },
    left: { style: 'double', size: 8 },
    bottom: { style: 'apples', size: 12 },
  });
  assert.deepEqual(section.lineNumbering, { countBy: 5, start: 2, distance: 360, restart: 'newSection' });
  // What was only partly read (an unknown child, the image of an art border) is kept.
  assert.deepEqual(section.extras?.map((extra) => [extra.el.name, extra.partial]), [
    ['w:endnotePr', true],
    ['w:pgBorders', true],
  ]);
});

// ------------------------------------------------------------- shown text

test('docx review: text moved away or deleted is not shown, alternate content shows its selected branch', async () => {
  const doc = await readSpec({
    body:
      '<w:p><w:r><w:t>a</w:t></w:r><w:moveFrom w:id="1" w:author="x"><w:r><w:t>moved</w:t></w:r></w:moveFrom><w:moveTo w:id="2" w:author="x"><w:r><w:t>b</w:t></w:r></w:moveTo>' +
      '<w:del w:id="3" w:author="x"><w:r><w:t>deleted</w:t><w:delText>too</w:delText></w:r><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:del>' +
      '<w:ins w:id="4" w:author="x"><w:hyperlink w:anchor="h"><w:r><w:t>c</w:t></w:r></w:hyperlink></w:ins>' +
      '<w:hyperlink w:anchor="h"><w:del w:id="5" w:author="x"><w:r><w:t>nested</w:t></w:r></w:del></w:hyperlink>' +
      '<w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:t>d</w:t><w:tab/></mc:Choice><mc:Fallback><w:t>fallback</w:t></mc:Fallback></mc:AlternateContent>' +
      '<mc:AlternateContent><mc:Choice Requires="nope"><w:t>choice</w:t></mc:Choice><mc:Fallback><w:t>e</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>',
  });
  assert.equal(paragraphText(para(doc.body.blocks, 0)), 'abcd\te');
});

test('docx review: only web, mail, phone and in-document hyperlink targets pass', () => {
  for (const [target, expected] of [
    ['https://example.com/a?b=1#c', 'https://example.com/a?b=1#c'],
    ['http://example.com', 'http://example.com'],
    ['HTTPS://EXAMPLE.COM', 'HTTPS://EXAMPLE.COM'],
    ['mailto:someone@example.com', 'mailto:someone@example.com'],
    ['tel:+900000000000', 'tel:+900000000000'],
    ['#bookmark', '#bookmark'],
    ['#', '#'],
    // Browsers ignore leading spaces and control characters, and tabs and line ends anywhere.
    ['  \u0000\u001f https://example.com', 'https://example.com'],
    ['ht\ttps://exam\nple.com/\r', 'https://example.com/'],
    ['https://exa\rmple.com/a\r\nb', 'https://example.com/ab'],
  ] as [string, string][]) {
    assert.equal(safeHyperlinkTarget(target), expected, JSON.stringify(target));
  }
  for (const target of [
    'javascript:alert(1)',
    'JavaScript:alert(1)',
    ' javascript:alert(1)',
    '\u0001javascript:alert(1)',
    'java\tscript:alert(1)',
    'java\r\nscript:alert(1)',
    'j\na\nv\na\ns\nc\nr\ni\np\nt\n:alert(1)',
    'data:text/html,<script>alert(1)</script>',
    'vbscript:msgbox(1)',
    'file:///etc/passwd',
    'FILE://server/share',
    'ftp://example.com',
    'blob:https://example.com/x',
    'about:blank',
    'ms-word:ofe|u|https://example.com/x.docx',
    'C:\\Users\\x\\file.docx',
    '\\\\server\\share\\file.docx',
    '//example.com/protocol-relative',
    'relative/path.html',
    '../up.html',
    'example.com',
    '?query',
    'https:',
    'http:/',
    'https:example.com',
    'mailto:',
    'tel:',
    ':nothing',
    '1http://example.com',
    '',
    '   ',
    undefined,
  ]) {
    assert.equal(safeHyperlinkTarget(target), undefined, JSON.stringify(target));
  }
});

// -------------------------------------------- behaviour no test pinned before

test('docx review: attributes of a properties element are kept with the properties', async () => {
  const doc = await readSpec({
    body:
      '<w:p><w:pPr w14:x="1"><w:jc w:val="center"/></w:pPr><w:r><w:rPr w14:y="2"><w:b/></w:rPr></w:r></w:p>' +
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr></w:r></w:p>',
  });
  const [withAttr, without] = [para(doc.body.blocks, 0), para(doc.body.blocks, 1)];
  assert.equal(withAttr.props.align, 'center');
  assert.deepEqual(withAttr.props.extras?.map((extra) => [extra.index, extra.el.name, extra.partial, extra.el.attrs.get('w14:x')]), [[-1, 'w:pPr', true, '1']]);
  assert.equal(without.props.extras, undefined);
  assert.deepEqual(run(withAttr, 0).props.extras?.map((extra) => [extra.index, extra.el.name, extra.partial]), [[-1, 'w:rPr', true]]);
  assert.equal(run(without, 0).props.extras, undefined);
});

test('docx review: properties are shared inside a part, never between parts', async () => {
  // The same source text means something else in another part: other relationships, other prefixes.
  const p = '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r></w:p>';
  const doc = await readSpec({ body: p + p, headers: { rIdH: p + p } });
  const [a, b] = [para(doc.body.blocks, 0), para(doc.body.blocks, 1)];
  const header = doc.headers.get('rIdH')!.blocks;
  const [c, d] = [para(header, 0), para(header, 1)];
  assert.equal(a.props, b.props);
  assert.equal(c.props, d.props);
  assert.notEqual(a.props, c.props);
  assert.deepEqual(a.props, c.props);
  assert.equal(run(a, 0).props, run(b, 0).props);
  assert.equal(run(c, 0).props, run(d, 0).props);
  assert.notEqual(run(a, 0).props, run(c, 0).props);
});

test('docx review: the first of several default styles of a type is the default', async () => {
  const doc = await readSpec({
    body: '<w:p/>',
    styles:
      '<w:style w:type="paragraph" w:styleId="A" w:default="1"/><w:style w:type="paragraph" w:styleId="B" w:default="1"/>' +
      '<w:style w:type="character" w:styleId="C" w:default="0"/><w:style w:type="character" w:styleId="D" w:default="true"/><w:style w:type="character" w:styleId="E" w:default="1"/>',
  });
  assert.deepEqual(doc.styles.defaultIds, { paragraph: 'A', character: 'D' });
  assert.deepEqual([...doc.styles.styles.values()].map((style) => style.isDefault), [true, true, false, true, true]);
});

// -------------------------------------------------------------- random damage

/** Small seeded generator, so that a failure can be reproduced. */
function random(seed: number): () => number {
  return () => {
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

test('docx review: a damaged document part gives an error or a model whose source spans are right', async () => {
  const rich =
    '<!--c--><w:p w:a=\'x>y\'><w:pPr><w:jc w:val="center"/><w:ind w:start="1"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/><w:instrText> PAGE </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/><w:t>1</w:t><w:fldChar w:fldCharType="end"/></w:r><w:hyperlink w:anchor="a"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>' +
    '<?pi x?><w:tbl><w:tblPr><w:tblLook w:val="04A0"/></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc>' +
    '<w:sdt><w:sdtContent><w:tc><w:p/></w:tc><w:bookmarkEnd w:id="1"/></w:sdtContent></w:sdt></w:tr></w:tbl>' +
    '<w:sdt><w:sdtContent><w:p><w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><a:graphic><a:graphicData uri="x/picture"><pic:pic><pic:blipFill><a:blip r:embed="r1"/></pic:blipFill></pic:pic>' +
    '<mc:AlternateContent><mc:Choice Requires="wps"><wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx></mc:Choice></mc:AlternateContent></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:sdtContent></w:sdt>' +
    '<![CDATA[ <w:p/> ]]><w:p><w:r><w:t xml:space="preserve"> &amp; &#x41; </w:t></w:r></w:p><w:sectPr><w:pgSz w:w="1" w:h="2"/></w:sectPr>';
  const good = documentXml(rich);
  const next = random(20261001);
  const pieces = ['<', '>', '/', '"', "'", '<!--', '-->', '<?', '?>', '<![CDATA[', ']]>', '&', ';', '</w:p>', '<w:p>', '<w:r>', ' ', 'w:', '=', '<w:fldChar w:fldCharType="begin"/>'];
  let read = 0;
  let refused = 0;
  for (let round = 0; round < 400; round++) {
    let text = good;
    const edits = 1 + Math.floor(next() * 3);
    for (let i = 0; i < edits; i++) {
      // Stay inside the body: a broken root element is covered elsewhere and always refused.
      const at = good.indexOf('<w:body>') + Math.floor(next() * (text.length - good.indexOf('<w:body>')));
      const kind = next();
      if (kind < 0.4) text = text.slice(0, at) + pieces[Math.floor(next() * pieces.length)] + text.slice(at);
      else if (kind < 0.8) text = text.slice(0, at) + text.slice(at + 1 + Math.floor(next() * 12));
      else text = text.slice(0, at);
    }
    let doc;
    try {
      doc = await readSpec({ document: text });
    } catch (error) {
      assert.ok(error instanceof Error && /^Invalid (XML|docx)|^Unknown XML entity/.test(error.message), `round ${round}: ${String(error)}`);
      refused++;
      continue;
    }
    read++;
    for (const block of doc.body.blocks) {
      const source = rawOf(doc.body, block.wrap ? outermost(block.wrap) : block);
      assert.ok(source !== undefined, `round ${round}: a ${block.type} block without a source span`);
      const expected = block.wrap ? '<' : block.type === 'paragraph' ? '<w:p' : block.type === 'table' ? '<w:tbl' : block.type === 'misc' ? (block.node.type === 'comment' ? '<!--' : '<?') : '<';
      assert.ok(source.startsWith(expected) && source.endsWith('>'), `round ${round}: the span of a ${block.type} block is "${source.slice(0, 20)}…"`);
    }
    for (const block of doc.body.blocks) if (block.type === 'paragraph') paragraphText(block);
  }
  // Both outcomes have to occur for the test to mean anything.
  assert.ok(read > 20 && refused > 20, `${read} read, ${refused} refused`);
});

function outermost(wrap: NonNullable<Block['wrap']>): NonNullable<Block['wrap']> {
  while (wrap.parent) wrap = wrap.parent;
  return wrap;
}
