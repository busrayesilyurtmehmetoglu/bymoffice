import assert from 'node:assert/strict';
import { test } from 'node:test';
import { DocLayout } from '../src/layout/doc/layout';
import { readSpec } from './docx-fixtures';
import { STYLES, at, baselines, fakeMeasurer, filler, find, lay, lines, p, pageText, r, round, sectPr, texts, words } from './doc-layout-fixtures';

// The standard page is 400 x 400 px with margins of 40 px: 20 lines of 16 px.

const field = (code: string, result: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve">${code}</w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  r(result) +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

/** A section mark inside a paragraph: ends a section that is not the last. */
const endSection = (inner = '', size = 'w:w="6000" w:h="6000"') =>
  `<w:p><w:pPr><w:sectPr>${inner}<w:pgSz ${size}/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr></w:pPr></w:p>`;

test('doc layout: page size, margins and the page count', async () => {
  const { pages } = await lay(filler(45));
  assert.equal(pages.length, 3);
  assert.deepEqual(pages.map((page) => [page.width, page.height, page.section, page.number, texts(page).length]), [
    [400, 400, 0, 1, 20],
    [400, 400, 0, 2, 20],
    [400, 400, 0, 3, 5],
  ]);
  assert.deepEqual(baselines(pages[1]).slice(0, 2), [52.8, 68.8]);
  assert.equal(round(baselines(pages[0])[19]), 356.8);

  // A document without a section mark gets US Letter with margins of an inch.
  const plain = new DocLayout(await readSpec({ body: p('x'), styles: STYLES }), fakeMeasurer);
  plain.layout();
  assert.deepEqual([plain.pages[0].width, plain.pages[0].height], [816, 1056]);
  assert.deepEqual(at(plain.pages[0]), ['x@96,108.8']);

  // Not even a paragraph: one empty page.
  const empty = new DocLayout(await readSpec({ body: '' }), fakeMeasurer);
  assert.deepEqual(empty.layout(), { done: true, updated: [] });
  assert.equal(empty.pages.length, 1);
  assert.deepEqual(empty.pages[0].items, []);
});

test('doc layout: sections start a new page, with their own paper', async () => {
  const landscape = '<w:sectPr><w:pgSz w:w="9000" w:h="6000" w:orient="landscape"/><w:pgMar w:top="300" w:right="300" w:bottom="300" w:left="900" w:header="150" w:footer="150" w:gutter="150"/></w:sectPr>';
  const { pages } = await lay(p('one') + endSection() + p(words(20)) + landscape);
  assert.equal(pages.length, 2);
  // The empty paragraph that carries the section mark is a line of the first section.
  assert.deepEqual(at(pages[0]), ['one@40,52.8']);
  assert.deepEqual([pages[1].width, pages[1].height, pages[1].section], [600, 400, 1]);
  // Left margin 60 plus gutter 10; 510 px of text: 63 characters a line.
  assert.deepEqual(at(pages[1]), ['www0 www1 www2 www3 www4 www5 www6 www7 www8 www9 ww10 ww11 @70,32.8', 'ww12 ww13 ww14 ww15 ww16 ww17 ww18 ww19@70,48.8']);
});

test('doc layout: a continuous section goes on below, in its own columns', async () => {
  const two = '<w:type w:val="continuous"/><w:cols w:num="2" w:space="720"/>';
  const { pages } = await lay(p('one') + p('two') + endSection() + filler(20) + sectPr(two));
  assert.equal(pages.length, 1);
  // Three lines of the first section, then columns of 136 px with 48 px between them, 17 lines high.
  const page = pages[0];
  assert.deepEqual(at(page).slice(0, 3), ['one@40,52.8', 'two@40,68.8', 'L0@40,100.8']);
  assert.equal(find(page, 'L16').y, 356.8);
  assert.deepEqual([find(page, 'L17').x, find(page, 'L17').y], [224, 100.8]);
  assert.equal(page.section, 0);
});

test('doc layout: columns that a continuous section break ends are balanced', async () => {
  const two = '<w:cols w:num="2" w:space="720"/>';
  // Ten lines and the empty paragraph of the section mark: six and five.
  const { pages } = await lay(filler(10) + endSection(two) + p('after') + sectPr('<w:type w:val="continuous"/>'));
  assert.equal(pages.length, 1);
  const page = pages[0];
  assert.deepEqual([find(page, 'L5').x, find(page, 'L5').y], [40, 132.8]);
  assert.deepEqual([find(page, 'L6').x, find(page, 'L6').y], [224, 52.8]);
  assert.deepEqual([find(page, 'L9').x, find(page, 'L9').y], [224, 100.8]);
  // The next section starts under the taller column.
  assert.deepEqual([find(page, 'after').x, find(page, 'after').y], [40, 148.8]);

  // More than a page of columns: full pages stay as they are, the last one is balanced.
  const long = await lay(filler(50) + endSection(two) + p('after') + sectPr('<w:type w:val="continuous"/>'));
  assert.equal(long.pages.length, 2);
  assert.equal(find(long.pages[0], 'L39').y, 356.8);
  assert.deepEqual([find(long.pages[1], 'L45').x, find(long.pages[1], 'L45').y], [40, 132.8]);
  assert.deepEqual([find(long.pages[1], 'L46').x, find(long.pages[1], 'L46').y], [224, 52.8]);
  assert.equal(find(long.pages[1], 'after').y, 148.8);
  // Lines that must stay together are not parted for the balance.
  const kept = await lay(p(words(60), '<w:keepLines/>') + endSection(two) + p('after') + sectPr('<w:type w:val="continuous"/>'));
  assert.equal(texts(kept.pages[0]).filter((item) => item.x === 224).length, 0);
});

test('doc layout: a continuous section has its own left and right margins', async () => {
  const narrow = '<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="6000" w:h="6000"/><w:pgMar w:top="600" w:right="1500" w:bottom="600" w:left="1500" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr>';
  const { pages } = await lay(p('one') + endSection() + p(words(8)) + narrow);
  // 100 px margins: 200 px of text, 25 characters.
  assert.deepEqual(at(pages[0]), ['one@40,52.8', 'www0 www1 www2 www3 www4 @100,84.8', 'www5 www6 www7@100,100.8']);
});

test('doc layout: a continuous section that changes the paper starts a page', async () => {
  const { pages } = await lay(p('one') + endSection() + p('two') + '<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="9000" w:h="6000"/></w:sectPr>');
  assert.deepEqual(pages.map((page) => page.width), [400, 600]);
});

test('doc layout: sections that start on an even or odd page leave a page empty when they must', async () => {
  const type = (value: string) => sectPr(`<w:type w:val="${value}"/>`);
  const even = await lay(p('one') + endSection() + p('two') + type('evenPage'));
  assert.deepEqual(even.pages.map((page) => [page.number, pageText(page)]), [
    [1, 'one'],
    [2, 'two'],
  ]);
  const odd = await lay(p('one') + endSection() + p('two') + type('oddPage'));
  assert.deepEqual(odd.pages.map((page) => [page.number, page.section, pageText(page)]), [
    [1, 0, 'one'],
    [2, 1, ''],
    [3, 1, 'two'],
  ]);
  assert.deepEqual(odd.pages[1].items, []);
});

test('doc layout: columns of equal and of stated widths; column breaks', async () => {
  const equal = await lay(filler(22) + sectPr('<w:cols w:num="2" w:space="300" w:sep="1"/>'));
  assert.equal(equal.pages.length, 1);
  // Two columns of 150 px, 20 px apart.
  assert.deepEqual([find(equal.pages[0], 'L19').x, find(equal.pages[0], 'L20').x, find(equal.pages[0], 'L20').y], [40, 210, 52.8]);
  // The line between the columns, as far down as the text goes.
  assert.deepEqual(
    lines(equal.pages[0]).map((line) => [line.x1, line.y1, line.x2, line.y2]),
    [[200, 40, 200, 360]],
  );
  const text = await lay(p(words(12)) + sectPr('<w:cols w:num="2" w:space="300"/>'));
  // 150 px: 18 characters a line.
  assert.equal(texts(text.pages[0])[0].text, 'www0 www1 www2 ');

  const stated = await lay(
    p(r('a') + '<w:r><w:br w:type="column"/></w:r>' + r(words(8))) + sectPr('<w:cols w:num="2" w:equalWidth="0"><w:col w:w="1500" w:space="300"/><w:col w:w="3000"/></w:cols>'),
  );
  // The second column starts at 40 + 100 + 20 and is 200 px wide: 25 characters.
  assert.deepEqual(at(stated.pages[0]), ['a@40,52.8', 'www0 www1 www2 www3 www4 @160,52.8', 'www5 www6 www7@160,68.8']);
});

test('doc layout: headers and footers; a tall one pushes the text', async () => {
  const refs = '<w:headerReference w:type="default" r:id="h1"/><w:footerReference w:type="default" r:id="f1"/>';
  const small = await lay({ body: filler(25) + sectPr(refs), headers: { h1: p('Header') }, footers: { f1: p('Footer') } });
  assert.equal(small.pages.length, 2);
  for (const page of small.pages) {
    // The header starts 10 px from the top, the footer ends 10 px from the bottom.
    assert.equal(find(page, 'Header').y, 22.8);
    assert.equal(find(page, 'Footer').y, 386.8);
    assert.equal(find(page, 'Header').source!.story, 'header:h1');
    assert.equal(find(page, 'Footer').source!.story, 'footer:f1');
  }
  assert.equal(find(small.pages[0], 'L0').y, 52.8);
  assert.equal(pageText(small.pages[1]).indexOf('L20'), 'Header'.length);

  const tall = await lay({ body: filler(25) + sectPr(refs), headers: { h1: filler(4, 'H') }, footers: { f1: filler(3, 'F') } });
  // Header: 10 + 64 = 74. Footer: 390 - 48 = 342. 16 lines fit between.
  assert.equal(find(tall.pages[0], 'L0').y, 86.8);
  assert.equal(find(tall.pages[0], 'F0').y, 354.8);
  assert.deepEqual(tall.pages.map((page) => texts(page).filter((item) => item.source!.story === 'body').length), [16, 9]);
});

test('doc layout: first page and even page headers; sections inherit the header of the section before', async () => {
  const refs = '<w:headerReference w:type="default" r:id="hd"/><w:headerReference w:type="first" r:id="hf"/><w:headerReference w:type="even" r:id="he"/><w:titlePg/>';
  const headers = { hd: p('Default'), hf: p('First'), he: p('Even') };
  const header = (page: { items: unknown[] }) =>
    texts(page as never)
      .filter((item) => item.source!.story.indexOf('header') === 0)
      .map((item) => item.text)
      .join('');
  const plain = await lay({ body: filler(70) + sectPr(refs), headers });
  assert.deepEqual(plain.pages.map(header), ['First', 'Default', 'Default', 'Default']);
  const evenOdd = await lay({ body: filler(70) + sectPr(refs), headers, settings: '<w:evenAndOddHeaders/>' });
  assert.deepEqual(evenOdd.pages.map(header), ['First', 'Even', 'Default', 'Even']);

  // The second section names no header: it uses those of the first, and has no first page of its own without titlePg.
  const inherited = await lay({ body: filler(25) + endSection(refs) + filler(25) + sectPr(), headers });
  assert.deepEqual(inherited.pages.map(header), ['First', 'Default', 'Default', 'Default']);
  // With titlePg but without a first header anywhere, the first page has none.
  const bare = await lay({ body: filler(25) + sectPr('<w:headerReference w:type="default" r:id="hd"/><w:titlePg/>'), headers: { hd: p('Default') } });
  assert.deepEqual(bare.pages.map(header), ['', 'Default']);
});

test('doc layout: page numbers in the footer, their format and restart', async () => {
  const footers = { f1: p(r('- ') + field(' PAGE ', '9') + r(' / ') + field(' NUMPAGES ', '9') + r(' -'), '<w:jc w:val="center"/>') };
  const ref = '<w:footerReference w:type="default" r:id="f1"/>';
  const footer = (page: { items: unknown[] }) =>
    texts(page as never)
      .filter((item) => item.source === undefined || item.source.story.indexOf('footer') === 0)
      .map((item) => item.text)
      .join('');
  const plain = await lay({ body: filler(45) + sectPr(ref), footers });
  assert.deepEqual(plain.pages.map(footer), ['- 1 / 3 -', '- 2 / 3 -', '- 3 / 3 -']);
  // Centred with its final text: 9 characters.
  assert.equal(find(plain.pages[0], '- ').x, 40 + (320 - 72) / 2);

  const roman = await lay({
    body: filler(25) + endSection(ref) + filler(25) + sectPr('<w:pgNumType w:fmt="lowerRoman" w:start="5"/>'),
    footers,
  });
  assert.deepEqual(roman.pages.map((page) => page.number), [1, 2, 5, 6]);
  assert.deepEqual(roman.pages.map(footer), ['- 1 / 4 -', '- 2 / 4 -', '- v / 4 -', '- vi / 4 -']);
});

test('doc layout: page fields inside tables, in the footer and in the text', async () => {
  const cell = (content: string) => `<w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
  const table = (inner: string) => `<w:tbl><w:tblPr><w:tblW w:w="4800" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid><w:tr>${inner}</w:tr></w:tbl>`;
  const footers = { f1: table(cell(p('Name')) + cell(p(r('Page ') + field(' PAGE ', '7') + r(' of ') + field(' NUMPAGES ', '7'), '<w:jc w:val="right"/>'))) + p('') };
  const body = filler(22) + table(cell(p(r('on ') + field(' PAGE ', '7'))) + cell(p('x'))) + sectPr('<w:footerReference w:type="default" r:id="f1"/>');
  const { pages } = await lay({ body, footers });
  assert.equal(pages.length, 2);
  const footer = (page: { items: unknown[] }) =>
    texts(page as never)
      .filter((item) => item.y > 360)
      .map((item) => item.text)
      .join('');
  assert.deepEqual(pages.map(footer), ['NamePage 1 of 2', 'NamePage 2 of 2']);
  // Right aligned in its cell with the final text: the cell ends at 40 - 7.2 + 320, less its margin.
  const last = texts(pages[0]).filter((item) => item.y > 360).pop()!;
  assert.equal(round(last.x + last.width), round(40 - 7.2 + 320 - 7.2));
  assert.equal(find(pages[1], 'on ').text + texts(pages[1])[texts(pages[1]).indexOf(find(pages[1], 'on ')) + 1].text, 'on 2');
});

test('doc layout: a page number in a text box of the footer', async () => {
  const box =
    '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>' +
    '<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
    '<wp:extent cx="952500" cy="476250"/><wp:wrapNone/><wp:docPr id="1" name="Box"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr/><wps:txbx><w:txbxContent>' +
    p(r('p') + field(' PAGE ', '7') + r('/') + field(' NUMPAGES ', '7')) +
    '</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
  const { pages } = await lay({ body: filler(45) + sectPr('<w:footerReference w:type="default" r:id="f1"/>'), footers: { f1: p(box) } });
  const boxed = (page: { items: unknown[] }) =>
    texts(page as never)
      .filter((item) => item.y < 30)
      .map((item) => item.text)
      .join('');
  assert.deepEqual(pages.map(boxed), ['p1/3', 'p2/3', 'p3/3']);
});

test('doc layout: page borders', async () => {
  const border = (side: string) => `<w:${side} w:val="double" w:sz="12" w:space="24" w:color="0000FF"/>`;
  const fromPage = await lay(p('x') + sectPr(`<w:pgBorders w:offsetFrom="page">${border('top')}${border('left')}${border('bottom')}${border('right')}</w:pgBorders>`));
  const drawn = lines(fromPage.pages[0]);
  assert.equal(drawn.length, 4);
  assert.ok(drawn.every((line) => line.color === '#0000FF' && line.style === 'double' && line.width === 2));
  // 24 pt = 32 px from the edge of the paper.
  assert.deepEqual(
    drawn.map((line) => [line.x1, line.y1, line.x2, line.y2]),
    [
      [32, 32, 368, 32],
      [32, 368, 368, 368],
      [32, 32, 32, 368],
      [368, 32, 368, 368],
    ],
  );
  const first = await lay(filler(25) + sectPr(`<w:pgBorders w:display="firstPage">${border('top')}</w:pgBorders>`));
  // From the text: 32 px above the top margin.
  assert.deepEqual(first.pages.map((page) => lines(page).map((line) => line.y1)), [[8], []]);
});

test('doc layout: mirrored margins swap on even pages', async () => {
  const margins = '<w:sectPr><w:pgSz w:w="6000" w:h="6000"/><w:pgMar w:top="600" w:right="300" w:bottom="600" w:left="900" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr>';
  const { pages } = await lay({ body: filler(45) + margins, settings: '<w:mirrorMargins/>' });
  assert.deepEqual(pages.map((page) => texts(page)[0].x), [60, 20, 60]);
});

test('doc layout: lines snap to the document grid when the section asks for it', async () => {
  const { pages } = await lay(p('one') + p('two') + sectPr('<w:docGrid w:type="lines" w:linePitch="360"/>'));
  // 24 px a line; the text sits in the middle of the spare room.
  assert.deepEqual(baselines(pages[0]), [56.8, 80.8]);
  const none = await lay(p('one') + p('two') + sectPr('<w:docGrid w:linePitch="360"/>'));
  assert.deepEqual(baselines(none.pages[0]), [52.8, 68.8]);
});

test('doc layout: layout in steps gives the same pages as layout at once', async () => {
  const footers = { f1: p(field(' PAGE ', '9') + r('/') + field(' NUMPAGES ', '9')) };
  const body =
    filler(30) +
    p(words(200), '<w:jc w:val="both"/><w:widowControl/>') +
    p('Heading', '<w:keepNext/>') +
    p(words(90)) +
    endSection('<w:footerReference w:type="default" r:id="f1"/>') +
    filler(60) +
    sectPr('<w:cols w:num="2"/>');
  const spec = { body, footers, styles: STYLES };
  const whole = new DocLayout(await readSpec(spec), fakeMeasurer);
  assert.deepEqual(whole.layout(), { done: true, updated: [] });
  assert.deepEqual(whole.layout(), { done: true, updated: [] });

  for (const step of [1, 3]) {
    const stepped = new DocLayout(await readSpec(spec), fakeMeasurer);
    let calls = 0;
    let result = stepped.layout({ pages: step });
    // The first pages are there before the rest is.
    assert.equal(stepped.pages.length, step);
    assert.equal(result.done, false);
    const early = JSON.stringify(stepped.pages[0]);
    while (!result.done) {
      assert.deepEqual(result.updated, []);
      result = stepped.layout({ pages: step });
      assert.ok(++calls < 100);
    }
    // Until the end the footer shows the cached total; then the pages that show it are named.
    assert.notEqual(early, JSON.stringify(stepped.pages[0]));
    assert.ok(result.updated.length > 0 && result.updated[0] === 0);
    // Only pages handed out before: the ones the last call added are not news.
    assert.ok(result.updated.every((index, k) => index === k) && result.updated.length < stepped.pages.length);
    assert.equal(JSON.stringify(stepped.pages), JSON.stringify(whole.pages));
  }

  const timed = new DocLayout(await readSpec(spec), fakeMeasurer);
  let calls = 0;
  while (!timed.layout({ milliseconds: 0 }).done) assert.ok(++calls < 100000);
  assert.ok(calls > 1);
  assert.equal(JSON.stringify(timed.pages), JSON.stringify(whole.pages));
});

test('doc layout: the page limit of the options', async () => {
  const doc = await readSpec({ body: filler(100) + sectPr(), styles: STYLES });
  const limited = new DocLayout(doc, fakeMeasurer, { maxPages: 2 });
  assert.deepEqual(limited.layout(), { done: false, updated: [] });
  assert.equal(limited.pages.length, 2);
  assert.deepEqual(limited.layout({ pages: 5 }), { done: false, updated: [] });
  assert.equal(limited.pages.length, 2);
});

test('doc layout: a font the host does not have is replaced by the fallback family', async () => {
  const measurer = { ...fakeMeasurer, hasFont: (family: string) => family === 'Known' };
  const { pages } = await lay(p(r('a', '<w:rFonts w:ascii="Known"/>') + r('b')), { fallbackFamily: 'Known Sans' }, measurer);
  assert.deepEqual(
    texts(pages[0]).map((item) => item.font.family),
    ['Known', 'Known Sans'],
  );
});
