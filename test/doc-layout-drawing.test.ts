import assert from 'node:assert/strict';
import { test } from 'node:test';
import type { ImageItem } from '../src/layout/doc/types';
import { CELL_MARGINS, COMPAT15, at, baselines, filler, find, lay, lines, p, pageText, r, round, tbl, tc, texts, tr, words } from './doc-layout-fixtures';

const PX = 9525;
const PICTURE =
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="x"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rIdImg"/><a:srcRect l="10000" t="20000"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm rot="5400000" flipH="1"><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic>';
const textBox = (inner: string) =>
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr/>' +
  `<wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>`;
const CHART = '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/></a:graphicData></a:graphic>';
const RELS: [string, string, string][] = [['rIdImg', 'image', 'media/image1.png']];

const inline = (width: number, height: number, graphic = PICTURE) =>
  `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${width * PX}" cy="${height * PX}"/><wp:effectExtent l="${PX}" t="0" r="${2 * PX}" b="0"/>` +
  `<wp:docPr id="1" name="Picture 1" descr="A picture"/>${graphic}</wp:inline></w:drawing></w:r>`;

interface AnchorSpec {
  width: number;
  height: number;
  h: string;
  v: string;
  wrap?: string;
  behind?: boolean;
  graphic?: string;
  dist?: number;
}
const offset = (px: number) => `<wp:posOffset>${px * PX}</wp:posOffset>`;
const align = (value: string) => `<wp:align>${value}</wp:align>`;
const anchor = (spec: AnchorSpec) =>
  `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="${(spec.dist || 0) * PX}" distR="${(spec.dist || 0) * PX}" simplePos="0" relativeHeight="1" behindDoc="${spec.behind ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1">` +
  `<wp:simplePos x="0" y="0"/>${spec.h}${spec.v}<wp:extent cx="${spec.width * PX}" cy="${spec.height * PX}"/>${spec.wrap || '<wp:wrapNone/>'}` +
  `<wp:docPr id="2" name="Anchored"/>${spec.graphic || PICTURE}</wp:anchor></w:drawing></w:r>`;
const posH = (from: string, inner: string) => `<wp:positionH relativeFrom="${from}">${inner}</wp:positionH>`;
const posV = (from: string, inner: string) => `<wp:positionV relativeFrom="${from}">${inner}</wp:positionV>`;

const images = (page: { items: unknown[] }) => (page.items as ImageItem[]).filter((item) => item.type === 'image');

test('doc layout: an inline picture is a box on the baseline that makes its line taller', async () => {
  const { pages } = await lay({ body: p(r('ab') + inline(100, 50) + r('cd')) + p('next'), rels: RELS });
  const [image] = images(pages[0]);
  // The line: 50 px above the baseline, the 3.2 px of descent below it.
  assert.deepEqual(at(pages[0]), ['ab@40,90', 'cd@156,90', 'next@40,106']);
  assert.deepEqual([image.x, image.y, image.width, image.height], [56, 40, 100, 50]);
  assert.equal(image.part, '/word/media/image1.png');
  assert.deepEqual(image.crop, { left: 0.1, top: 0.2, right: 0, bottom: 0 });
  assert.equal(image.rotation, 90);
  assert.equal(image.flipH, true);
  assert.equal(image.alt, 'A picture');
  assert.deepEqual(image.effect, { left: 1, top: 0, right: 2, bottom: 0 });
});

test('doc layout: a picture wider than what is left of the line goes to the next line', async () => {
  const { pages } = await lay({ body: p(r(words(6)) + inline(100, 20) + inline(300, 20) + inline(500, 20)), rels: RELS });
  // 30 characters end at 280: 100 px do not fit behind them. The widest one is placed although it does not fit anywhere.
  assert.deepEqual(
    images(pages[0]).map((image) => [image.x, image.y]),
    [
      [40, 56],
      [40, 79.2],
      [40, 102.4],
    ],
  );
});

test('doc layout: a picture taller than the page is placed and layout goes on', async () => {
  const { pages } = await lay({ body: p('one') + p(inline(100, 1000)) + p('two'), rels: RELS });
  assert.equal(pages.length, 3);
  assert.equal(images(pages[1]).length, 1);
  assert.deepEqual(at(pages[2]), ['two@40,52.8']);
});

test('doc layout: anchored drawings are placed relative to the page, the margins, the column or the paragraph', async () => {
  const body =
    p('first') +
    p(
      r('second') +
        anchor({ width: 100, height: 50, h: posH('column', offset(20)), v: posV('paragraph', offset(10)) }) +
        anchor({ width: 100, height: 50, h: posH('page', align('center')), v: posV('margin', align('bottom')) }) +
        anchor({ width: 100, height: 50, h: posH('margin', align('right')), v: posV('page', offset(5)), behind: true }) +
        anchor({ width: 100, height: 50, h: posH('rightMargin', offset(0)), v: posV('topMargin', align('center')) }),
    );
  const { pages } = await lay({ body, rels: RELS });
  const page = pages[0];
  // The text is not moved by drawings without wrapping.
  assert.deepEqual(at(page), ['first@40,52.8', 'second@40,68.8']);
  assert.deepEqual(
    images(page).map((image) => [image.x, image.y]),
    [
      // Behind the text: painted first.
      [260, 5],
      // 20 px into the column, 10 px below the top of the second paragraph.
      [60, 66],
      [150, 310],
      [360, -5],
    ],
  );
  assert.ok(page.items.indexOf(images(page)[0]) < page.items.indexOf(find(page, 'first')));
  assert.ok(page.items.indexOf(images(page)[1]) > page.items.indexOf(find(page, 'second')));
});

test('doc layout: text moves below a drawing with top and bottom wrapping', async () => {
  const drawing = anchor({ width: 100, height: 50, h: posH('column', offset(0)), v: posV('paragraph', offset(0)), wrap: '<wp:wrapTopAndBottom/>' });
  const { pages } = await lay({ body: p('before') + p(drawing + r(words(12))) + p('after'), rels: RELS });
  assert.deepEqual([images(pages[0])[0].x, images(pages[0])[0].y], [40, 56]);
  assert.deepEqual(baselines(pages[0]), [52.8, 118.8, 134.8, 150.8]);
});

test('doc layout: text runs beside a drawing with square wrapping, and at full width below it', async () => {
  const drawing = anchor({ width: 100, height: 50, h: posH('column', offset(0)), v: posV('paragraph', offset(0)), wrap: '<wp:wrapSquare wrapText="bothSides"/>', dist: 10 });
  const { pages } = await lay({ body: p(drawing + r(words(40))) + p(words(10)), rels: RELS });
  const items = texts(pages[0]);
  // The drawing and its distance take 110 px: 210 px (26 characters) are left for the four lines beside it.
  assert.deepEqual(
    items.slice(0, 6).map((item) => [item.x, item.y, item.text.length]),
    [
      [150, 52.8, 25],
      [150, 68.8, 25],
      [150, 84.8, 25],
      [150, 100.8, 25],
      [40, 116.8, 40],
      [40, 132.8, 40],
    ],
  );
  assert.equal(pages[0].items.filter((item) => item.type === 'text' && item.source).map((item) => (item as { text: string }).text).join(''), words(40) + words(10));

  // A drawing on the right: the text keeps to the left of it.
  const right = anchor({ width: 100, height: 50, h: posH('column', align('right')), v: posV('paragraph', offset(0)), wrap: '<wp:wrapSquare wrapText="bothSides"/>' });
  const other = await lay({ body: p(right + r(words(20))), rels: RELS });
  assert.deepEqual(
    texts(other.pages[0]).map((item) => [item.x, item.text.length]),
    [
      [40, 25],
      [40, 25],
      [40, 25],
      [40, 24],
    ],
  );
});

test('doc layout: a paragraph beside a drawing that moves to the next page is broken again there', async () => {
  const drawing = anchor({ width: 100, height: 50, h: posH('column', offset(0)), v: posV('paragraph', offset(0)), wrap: '<wp:wrapSquare wrapText="bothSides"/>' });
  const { pages } = await lay({ body: filler(19) + p(drawing + r(words(40)), '<w:keepLines/>'), rels: RELS });
  assert.equal(pages.length, 2);
  assert.equal(images(pages[0]).length, 0);
  assert.deepEqual([images(pages[1])[0].x, images(pages[1])[0].y], [40, 40]);
  assert.deepEqual(
    texts(pages[1]).slice(0, 5).map((item) => [item.x, item.y]),
    [
      [140, 52.8],
      [140, 68.8],
      [140, 84.8],
      [140, 100.8],
      [40, 116.8],
    ],
  );
  assert.equal(pageText(pages[1]), words(40));
});

test('doc layout: a drawing that leaves no room on its page does not push the text off it', async () => {
  const drawing = anchor({ width: 320, height: 330, h: posH('column', offset(0)), v: posV('paragraph', offset(0)), wrap: '<wp:wrapTopAndBottom/>' });
  const { pages } = await lay({ body: p(drawing + r(words(16))) + p('after'), rels: RELS });
  // The first line stays with the drawing it anchors (and runs over it); the rest goes to the next page.
  assert.equal(pages.length, 2);
  assert.equal(images(pages[0]).length, 1);
  assert.deepEqual(at(pages[0]), ['www0 www1 www2 www3 www4 www5 www6 www7 @40,52.8']);
  assert.deepEqual(at(pages[1]), ['www8 www9 ww10 ww11 ww12 ww13 ww14 ww15@40,52.8', 'after@40,68.8']);
});

test('doc layout: the text of a text box is laid out inside the box', async () => {
  const box = anchor({ width: 200, height: 100, h: posH('column', offset(50)), v: posV('paragraph', offset(20)), graphic: textBox(p('inside the box') + p(words(8))) });
  const { pages } = await lay(p(r('host') + box));
  // Insets of 0.1 and 0.05 inch; 180.8 px of width hold 22 characters.
  assert.deepEqual(at(pages[0]), ['host@40,52.8', 'inside the box@99.6,77.6', 'www0 www1 www2 www3 @99.6,93.6', 'www4 www5 www6 www7@99.6,109.6']);
  assert.deepEqual(find(pages[0], 'inside').source, { story: 'body', path: [0, -1, 0, 0, 0], offset: 0 });
  // Insets the shape states itself.
  const tight = textBox(p('tight')).replace('<wps:bodyPr/>', '<wps:bodyPr lIns="0" tIns="19050" rIns="0" bIns="0"/>');
  const stated = await lay(p(r('host') + anchor({ width: 200, height: 100, h: posH('column', offset(50)), v: posV('paragraph', offset(20)), graphic: tight })));
  assert.deepEqual(at(stated.pages[0]), ['host@40,52.8', 'tight@90,74.8']);
  // The same inline: the box is part of the line.
  const inlined = await lay(p(r('a') + inline(200, 100, textBox(p('boxed'))) + r('b')));
  assert.deepEqual(at(inlined.pages[0]), ['a@40,140', 'boxed@57.6,57.6', 'b@248,140']);
});

test('doc layout: charts and legacy graphics leave a placeholder of their size', async () => {
  const pict = '<w:r><w:pict><v:shape id="s1" style="width:75pt;height:37.5pt"><v:textpath string="x"/></v:shape></w:pict></w:r>';
  const floating = '<w:r><w:pict><v:rect style="position:absolute;margin-left:30pt;margin-top:15pt;width:75pt;height:37.5pt;z-index:-1"/></w:pict></w:r>';
  const { pages } = await lay(p(r('a') + inline(100, 50, CHART)) + p(r('b') + pict) + p(r('c') + floating));
  const found = images(pages[0]).map((image) => [image.placeholder, image.x, round(image.y), image.width, image.height]);
  assert.deepEqual(found, [
    // Absolutely positioned and behind the text: 40 and 20 px from the paragraph.
    ['vml', 80, round(40 + 53.2 + 53.2 + 20), 100, 50],
    ['chart', 48, 40, 100, 50],
    ['vml', 48, 93.2, 100, 50],
  ]);
});

test('doc layout: drawings in table cells are placed relative to the cell', async () => {
  const drawing = anchor({ width: 20, height: 20, h: posH('column', offset(5)), v: posV('paragraph', offset(3)) });
  const { pages } = await lay({ body: tbl(tr(tc('a') + tc(p('one') + p(r('two') + drawing))), [1500, 3000], CELL_MARGINS), settings: COMPAT15, rels: RELS });
  assert.deepEqual([images(pages[0])[0].x, images(pages[0])[0].y], [155, 59]);
});

const NOTES =
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Note one</w:t></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Note two</w:t></w:r></w:p></w:footnote>';
const ref = (id: number) => `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;

test('doc layout: footnotes sit at the foot of the page under a separator and take room from the text', async () => {
  const { pages } = await lay({ body: p(r('Text') + ref(1) + r(' more') + ref(2)) + filler(25), footnotes: NOTES });
  const page = pages[0];
  const items = texts(page);
  // The marks: small raised numbers that have no source. They make their line 0.8 px taller.
  assert.deepEqual(
    items.slice(0, 4).map((item) => [item.text, round(item.x), round(item.y), round(item.font.size)]),
    [
      ['Text', 40, 53.6, 16],
      ['1', 72, 48.32, 10.4],
      [' more', 77.2, 53.6, 16],
      ['2', 117.2, 48.32, 10.4],
    ],
  );
  assert.equal(items[1].source, undefined);
  // Two notes of 16.8 px and the separator: the text has 320 - 43.6 px, the marked line and 16 more.
  const body = items.filter((item) => item.source && item.source.story === 'body');
  assert.equal(body[body.length - 1].text, 'L15');
  const [separator] = lines(page);
  assert.deepEqual([separator.x1, round(separator.y1), separator.x2], [40, round(360 - 43.6 + 5), 232]);
  const notes = items.filter((item) => item.source && item.source.story.indexOf('footnote') === 0);
  assert.deepEqual(
    notes.map((item) => [item.text, item.source!.story, round(item.x), round(item.y)]),
    [
      [' Note one', 'footnote:1', 45.2, round(360 - 33.6 + 13.6)],
      [' Note two', 'footnote:2', 45.2, round(360 - 16.8 + 13.6)],
    ],
  );
  // The numbers in front of the notes.
  assert.deepEqual(
    items.filter((item) => !item.source && round(item.x) === 40).map((item) => item.text),
    ['1', '2'],
  );
  assert.equal(lines(pages[1]).length, 0);
});

test('doc layout: a line goes to the next page with its footnote when the two do not fit', async () => {
  const { pages } = await lay({ body: filler(19) + p(r('Text') + ref(1)), footnotes: NOTES });
  assert.equal(pages.length, 2);
  assert.equal(texts(pages[0]).length, 19);
  assert.equal(lines(pages[0]).length, 0);
  assert.equal(pageText(pages[1]), 'Text Note one');
});

test('doc layout: a footnote too long for its page goes on at the foot of the next', async () => {
  let long = '<w:footnote w:id="1"><w:p><w:r><w:t>N0</w:t></w:r></w:p>';
  for (let i = 1; i < 30; i++) long += p('N' + i);
  long += '</w:footnote>';
  const { pages } = await lay({ body: p(r('Text') + ref(1)) + p('next'), footnotes: long });
  assert.equal(pages.length, 2);
  const count = (page: { items: unknown[] }) => texts(page as never).filter((item) => item.source && item.source.story === 'footnote:1').length;
  // The note takes what is left of the page of its reference (293.2 px less the separator: 18 lines);
  // the text after the reference goes to the next page, where the rest of the note is.
  assert.equal(count(pages[0]) + count(pages[1]), 30);
  assert.equal(count(pages[0]), 18);
  assert.equal(pageText(pages[0]).indexOf('TextN0'), 0);
  assert.equal(pageText(pages[1]).indexOf('next'), 0);
  assert.equal(lines(pages[1]).length, 1);
  assert.equal(find(pages[1], 'N29').y, 356.8);
});

test('doc layout: a footnote referred to in a table cell', async () => {
  const { pages } = await lay({ body: tbl(tr(tc(p(r('Cell') + ref(1))) + tc('other')), [1500, 1500], CELL_MARGINS) + p('after'), footnotes: NOTES, settings: COMPAT15 });
  assert.equal(pages.length, 1);
  assert.equal(pageText(pages[0]), 'Cellotherafter Note one');
  assert.equal(round(find(pages[0], ' Note one').y), 356.8);
});

test('doc layout: endnotes follow the text of the document', async () => {
  const endnotes = '<w:endnote w:id="1"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t xml:space="preserve"> End one</w:t></w:r></w:p></w:endnote>';
  const { pages } = await lay({ body: p(r('Text') + '<w:r><w:endnoteReference w:id="1"/></w:r>') + p('last'), endnotes });
  const items = texts(pages[0]);
  assert.deepEqual(
    items.map((item) => item.text),
    ['Text', 'i', 'last', 'i', ' End one'],
  );
  assert.equal(items[4].source!.story, 'endnote:1');
  assert.equal(lines(pages[0]).length, 1);
  assert.ok(items[4].y > items[2].y + 16);
});
