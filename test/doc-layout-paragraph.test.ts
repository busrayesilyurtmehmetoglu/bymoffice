import assert from 'node:assert/strict';
import { test } from 'node:test';
import { at, baselines, filler, find, lay, lines, p, r, rects, round, texts, words } from './doc-layout-fixtures';

// A character is 8 px, a line 16 px (ascent 12.8); the text area is x 40..360, y 40..360.

test('doc layout: paragraphs follow each other, lines wrap at the right margin', async () => {
  const { pages } = await lay(p('Hello world') + p(words(20)));
  assert.equal(pages.length, 1);
  assert.equal(pages[0].width, 400);
  assert.equal(pages[0].height, 400);
  assert.deepEqual(at(pages[0]), [
    'Hello world@40,52.8',
    'www0 www1 www2 www3 www4 www5 www6 www7 @40,68.8',
    'www8 www9 ww10 ww11 ww12 ww13 ww14 ww15 @40,84.8',
    'ww16 ww17 ww18 ww19@40,100.8',
  ]);
  // The space a line ends with hangs out of the line: it does not count for the width.
  assert.equal(texts(pages[0])[1].width, 39 * 8);
});

test('doc layout: indents (left, right, first line, hanging)', async () => {
  const text = words(16);
  const { pages } = await lay(p(text, '<w:ind w:left="300" w:right="150" w:firstLine="150"/>') + p(text, '<w:ind w:left="300" w:hanging="150"/>'));
  const items = texts(pages[0]);
  // First paragraph: 280 px on the first line (7 words), 290 px after it.
  assert.equal(items[0].x, 70);
  assert.equal(items[0].text, 'www0 www1 www2 www3 www4 www5 www6 ');
  assert.equal(items[1].x, 60);
  assert.equal(items[1].text, 'www7 www8 www9 ww10 ww11 ww12 ww13 ');
  // Second paragraph: the first line starts left of the others.
  const second = items.filter((item) => item.source!.path[0] === 1);
  assert.equal(second[0].x, 50);
  assert.equal(second[1].x, 60);
});

test('doc layout: space before and after add up; contextual spacing removes them between like paragraphs', async () => {
  const spaced = '<w:spacing w:before="150" w:after="300"/>';
  const a = await lay(p('one', spaced) + p('two', spaced) + p('three'));
  // 10 before, 16 line, 20 after + 10 before, ...
  assert.deepEqual(baselines(a.pages[0]), [62.8, 108.8, 144.8]);

  const styles =
    '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:styleId="A"><w:name w:val="A"/><w:pPr><w:spacing w:before="150" w:after="300"/><w:contextualSpacing/></w:pPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="B"><w:name w:val="B"/></w:style>';
  const style = (id: string) => `<w:pStyle w:val="${id}"/>`;
  const b = await lay({ body: p('one', style('A')) + p('two', style('A')) + p('three', style('B')) + p('four', style('A')), styles });
  // No space between the two A paragraphs; the usual space where the style changes.
  assert.deepEqual(baselines(b.pages[0]), [62.8, 78.8, 114.8, 140.8]);
});

test('doc layout: automatic spacing merges like HTML margins, and not between items of a list', async () => {
  const auto = '<w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1"/>';
  const a = await lay(p('one', auto) + p('two', auto) + p('three', '<w:spacing w:before="600"/>'));
  // 14 pt = 18.67 px before the first, the same once between the two, then the larger of 18.67 and 40.
  assert.deepEqual(baselines(a.pages[0]), [round(52.8 + 18.667), round(52.8 + 18.667 * 2 + 16), round(52.8 + 18.667 * 2 + 16 + 40 + 16)]);

  const settings = '<w:compat><w:doNotUseHTMLParagraphAutoSpacing/></w:compat>';
  const b = await lay({ body: p('one', auto) + p('two', auto), settings });
  // With the compatibility setting the stated 5 points apply, and add up.
  assert.deepEqual(baselines(b.pages[0]), [round(52.8 + 6.667), round(52.8 + 6.667 * 3 + 16)]);
});

test('doc layout: space before is dropped at the top of a page after an automatic break, kept after a page break', async () => {
  const before = '<w:spacing w:before="300"/>';
  const a = await lay(filler(20) + p('next', before));
  assert.equal(a.pages.length, 2);
  assert.deepEqual(baselines(a.pages[1]), [52.8]);
  const b = await lay(filler(3) + p('next', before + '<w:pageBreakBefore/>'));
  assert.equal(b.pages.length, 2);
  assert.deepEqual(baselines(b.pages[1]), [72.8]);
  // The first paragraph of the document keeps it too.
  const c = await lay(p('first', before));
  assert.deepEqual(baselines(c.pages[0]), [72.8]);
});

test('doc layout: line spacing auto, exact and at least', async () => {
  const text = words(16);
  const double = await lay(p(text, '<w:spacing w:line="480" w:lineRule="auto"/>') + p('after'));
  // Twice the natural height; the extra room is below the text.
  assert.deepEqual(baselines(double.pages[0]), [52.8, 84.8, 116.8]);
  const exact = await lay(p(text, '<w:spacing w:line="300" w:lineRule="exact"/>') + p('after'));
  // 20 px a line, the text at the bottom of it.
  assert.deepEqual(baselines(exact.pages[0]), [56.8, 76.8, 92.8]);
  const tight = await lay(p(text, '<w:spacing w:line="150" w:lineRule="exact"/>') + p('after'));
  assert.deepEqual(baselines(tight.pages[0]), [46.8, 56.8, 72.8]);
  const atLeast = await lay(p(text, '<w:spacing w:line="300" w:lineRule="atLeast"/>') + p(text, '<w:spacing w:line="150" w:lineRule="atLeast"/>'));
  assert.deepEqual(baselines(atLeast.pages[0]), [56.8, 76.8, 92.8, 108.8]);
});

test('doc layout: alignment left, center, right and justified', async () => {
  const { pages } = await lay(p('abcd', '<w:jc w:val="center"/>') + p('abcd', '<w:jc w:val="right"/>') + p(words(12), '<w:jc w:val="both"/>'));
  const items = texts(pages[0]);
  assert.equal(items[0].x, 40 + (320 - 32) / 2);
  assert.equal(items[1].x, 360 - 32);
  // Justified: the 8 spare px of the first line go into its 7 spaces; the last line is not stretched.
  const first = items.filter((item) => round(item.y) === 84.8);
  assert.equal(first.length, 8);
  first.forEach((item, k) => assert.equal(round(item.x), round(40 + k * (40 + 8 / 7))));
  assert.equal(round(first[7].x + first[7].width), 360);
  const last = items.filter((item) => round(item.y) === 100.8);
  assert.deepEqual(
    last.map((item) => [item.text, item.x]),
    [['www8 www9 ww10 ww11', 40]],
  );
});

test('doc layout: a justified line that ends with a line break is stretched unless the document says not to', async () => {
  const body = p(r('aaaa bbbb cccc') + '<w:r><w:br/></w:r>' + r('dddd'), '<w:jc w:val="both"/>');
  const a = await lay(body);
  assert.equal(round(find(a.pages[0], 'cccc').x + find(a.pages[0], 'cccc').width), 360);
  const b = await lay({ body, settings: '<w:compat><w:doNotExpandShiftReturn/></w:compat>' });
  assert.equal(find(b.pages[0], 'aaaa').text, 'aaaa bbbb cccc');
});

test('doc layout: tab stops (default, left, center, right, decimal, leader)', async () => {
  const tab = (text: string) => `<w:r><w:tab/><w:t>${text}</w:t></w:r>`;
  const tabs = (inner: string) => `<w:tabs>${inner}</w:tabs>`;
  const { pages } = await lay(
    p(r('a') + tab('b') + tab('c')) +
      p(r('a') + tab('b'), tabs('<w:tab w:val="left" w:pos="1500"/>')) +
      p(r('a') + tab('bbb'), tabs('<w:tab w:val="right" w:pos="3000"/>')) +
      p(r('a') + tab('bbbb'), tabs('<w:tab w:val="center" w:pos="3000"/>')) +
      p(r('a') + tab('12.5'), tabs('<w:tab w:val="decimal" w:pos="3000"/>')) +
      p(r('a') + tab('b'), tabs('<w:tab w:val="right" w:leader="dot" w:pos="4500"/>')) +
      p(r('a') + tab('b') + tab('c'), tabs('<w:tab w:val="left" w:pos="1500"/>')),
  );
  assert.deepEqual(at(pages[0]), [
    // Default stops every 48 px.
    'a@40,52.8',
    'b@88,52.8',
    'c@136,52.8',
    'a@40,68.8',
    'b@140,68.8',
    'a@40,84.8',
    'bbb@216,84.8',
    'a@40,100.8',
    'bbbb@224,100.8',
    // The decimal point sits at the stop.
    'a@40,116.8',
    '12.5@224,116.8',
    // 35 dots fill the 284 px of the tab, ending at the text.
    'a@40,132.8',
    '...................................@52,132.8',
    'b@332,132.8',
    // After the last custom stop the default stops go on: 100 px, then the next multiple of 48.
    'a@40,148.8',
    'b@140,148.8',
    'c@184,148.8',
  ]);
});

test('doc layout: a hanging indent is a tab stop; tabs that pass the right edge wrap', async () => {
  const { pages } = await lay(
    p(r('1.') + '<w:r><w:tab/><w:t>text</w:t></w:r>', '<w:ind w:left="600" w:hanging="600"/>') + p(r(words(7)) + '<w:r><w:tab/><w:tab/><w:t>end</w:t></w:r>'),
  );
  assert.deepEqual(at(pages[0]).slice(0, 2), ['1.@40,52.8', 'text@80,52.8']);
  // 34 characters end at 312; the first tab runs to 328 (stop 288 px), the next default stop is past the edge.
  assert.equal(find(pages[0], 'end').y, 84.8);
  assert.equal(find(pages[0], 'end').x, 40 + 48);
});

const NUMBERING =
  '<w:abstractNum w:abstractNumId="0">' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/><w:lvlJc w:val="right"/><w:suff w:val="space"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr><w:rPr><w:b/></w:rPr></w:lvl>' +
  '</w:abstractNum>' +
  '<w:abstractNum w:abstractNumId="1">' +
  '<w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0A7;"/><w:lvlJc w:val="left"/><w:suff w:val="nothing"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr></w:lvl>' +
  '<w:lvl w:ilvl="2"><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr></w:lvl>' +
  '</w:abstractNum>' +
  '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>';
const num = (id: number, level = 0) => `<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${id}"/></w:numPr>`;

test('doc layout: numbers and bullets of lists', async () => {
  const { pages } = await lay({
    body: p('one', num(1)) + p('two', num(1)) + p('sub', num(1, 1)) + p('dot', num(2)) + p('square', num(2, 1)) + p('circle', num(2, 2)) + p(words(12), num(1)),
    numbering: NUMBERING,
  });
  const items = texts(pages[0]);
  assert.deepEqual(at(pages[0]).slice(0, 11), [
    // The number starts at the hanging indent (40 + 48 - 24), the text at the indent.
    '1.@64,52.8',
    'one@88,52.8',
    '2.@64,68.8',
    'two@88,68.8',
    // Right aligned: the number ends where the first line starts; bold, 4 characters of 8.8 px; then a (bold) space.
    '2.a)@76.8,84.8',
    'sub@120.8,84.8',
    '•@64,100.8',
    'dot@88,100.8',
    // No suffix: the text follows the bullet at once.
    '▪@64,116.8',
    'square@72,116.8',
    'o@64,132.8',
  ]);
  // The label has no source and takes the formatting of the level.
  assert.equal(items[0].source, undefined);
  assert.equal(items[4].font.bold, true);
  assert.equal(items.filter((item) => item.text === 'o')[0].font.family, 'Courier New');
  // Mapped symbols are ordinary characters: they do not ask for the symbol font.
  assert.notEqual(find(pages[0], '•').font.family, 'Symbol');
  // Wrapped lines of a list item start at the indent.
  assert.deepEqual(at(pages[0]).slice(12), ['3.@64,148.8', 'www0 www1 www2 www3 www4 www5 www6 @88,148.8', 'www7 www8 www9 ww10 ww11@88,164.8']);
});

test('doc layout: keep with next moves a heading to the page of the text it belongs to', async () => {
  const { pages } = await lay(filler(19) + p('Heading', '<w:keepNext/>') + p(words(16)));
  assert.equal(pages.length, 2);
  assert.equal(texts(pages[0]).length, 19);
  assert.deepEqual(at(pages[1]).slice(0, 2), ['Heading@40,52.8', 'www0 www1 www2 www3 www4 www5 www6 www7 @40,68.8']);
  // Without it the heading stays at the foot of the first page.
  const loose = await lay(filler(19) + p('Heading') + p(words(16)));
  assert.equal(texts(loose.pages[0]).length, 20);
});

test('doc layout: a chain of keep-with-next paragraphs moves as a whole, and gives up when it cannot fit', async () => {
  const keep = '<w:keepNext/>';
  const { pages } = await lay(filler(17) + p('A', keep) + p('B', keep) + p('C', keep) + p('D'));
  assert.equal(texts(pages[0]).length, 17);
  assert.deepEqual(at(pages[1]), ['A@40,52.8', 'B@40,68.8', 'C@40,84.8', 'D@40,100.8']);
  // Every paragraph of the document keeps with the next: pages are simply filled.
  let all = '';
  for (let i = 0; i < 50; i++) all += p('K' + i, keep);
  const chained = await lay(all);
  assert.equal(chained.pages.length, 3);
  assert.deepEqual(chained.pages.map((page) => texts(page).length), [20, 20, 10]);
});

test('doc layout: keep lines together, widow and orphan control', async () => {
  const text = words(32); // four lines
  const kept = await lay(filler(18) + p(text, '<w:keepLines/>'));
  assert.deepEqual(kept.pages.map((page) => texts(page).length), [18, 4]);
  const loose = await lay(filler(18) + p(text));
  assert.deepEqual(loose.pages.map((page) => texts(page).length), [20, 2]);

  const widow = '<w:widowControl/>';
  // One line would fit: an orphan. The paragraph can only be divided two and two, so it moves.
  const orphan = await lay(filler(19) + p(text, widow));
  assert.deepEqual(orphan.pages.map((page) => texts(page).length), [19, 4]);
  // Three lines would fit and leave a widow: two stay, two go.
  const two = await lay(filler(17) + p(text, widow));
  assert.deepEqual(two.pages.map((page) => texts(page).length), [19, 2]);
  // Three lines cannot be divided at all.
  const three = await lay(filler(18) + p(words(24), widow));
  assert.deepEqual(three.pages.map((page) => texts(page).length), [18, 3]);
});

test('doc layout: paragraph borders and shading', async () => {
  const border = (side: string) => `<w:${side} w:val="single" w:sz="8" w:space="3" w:color="FF0000"/>`;
  const boxed = `<w:pBdr>${border('top')}${border('left')}${border('bottom')}${border('right')}</w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>`;
  const { pages } = await lay(p('boxed', boxed) + p('after'));
  const page = pages[0];
  // Each border takes 1 pt of line and 3 pt of space: 5.33 px above and below the text.
  const room = 4 / 3 + 4;
  assert.deepEqual(baselines(page), [round(52.8 + room), round(52.8 + room * 2 + 16)]);
  const [shade] = rects(page);
  assert.deepEqual([shade.fill, round(shade.x), shade.y, round(shade.width), round(shade.height)], ['#FFFF00', 36, 40, 328, round(16 + 2 * room)]);
  const drawn = lines(page);
  assert.equal(drawn.length, 4);
  assert.ok(drawn.every((line) => line.color === '#FF0000' && round(line.width) === 1.33 && line.style === 'single'));
  const horizontal = drawn.filter((line) => line.y1 === line.y2).map((line) => [round(line.x1), round(line.x2), round(line.y1)]);
  assert.deepEqual(horizontal, [
    [36, 364, round(40 + 2 / 3)],
    [36, 364, round(40 + 16 + 2 * room - 2 / 3)],
  ]);
});

test('doc layout: neighbours with the same borders share one box', async () => {
  const border = (side: string) => `<w:${side} w:val="single" w:sz="8" w:space="0" w:color="000000"/>`;
  const boxed = `<w:pBdr>${border('top')}${border('bottom')}${border('between')}</w:pBdr>`;
  const { pages } = await lay(p('one', boxed) + p('two', boxed) + p('three'));
  const horizontal = lines(pages[0]).map((line) => round(line.y1));
  // Top of the box, the line between the two, the bottom: three lines instead of four.
  assert.equal(horizontal.length, 3);
  const w = 4 / 3;
  assert.deepEqual(horizontal, [round(40 + w / 2), round(40 + w + 16 + w / 2), round(40 + w + 16 + w + 16 + w / 2)]);
});
