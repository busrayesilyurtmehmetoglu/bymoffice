// Tests for defects found by the independent review of the document layout
// engine: each of them failed before its fix.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { getPrep, layoutBlocks, paragraphEnv } from '../src/layout/doc/blocks';
import { LayoutContext } from '../src/layout/doc/context';
import { DocLayout } from '../src/layout/doc/layout';
import { paragraphBands } from '../src/layout/doc/paragraph';
import type { FontMetrics, ImageItem, LayoutFont, LayoutPage, TextItem, TextMeasurer } from '../src/layout/doc/types';
import { MAX_PARAGRAPH_LINES } from '../src/layout/doc/units';
import type { Paragraph } from '../src/model/doc/types';
import { CELL_MARGINS, STYLES, at, baselines, fakeMeasurer, find, lay, lines, p, pageText, r, rects, round, sectPr, tbl, tc, texts, tr, words } from './doc-layout-fixtures';
import { THEME, readSpec } from './docx-fixtures';

const PX = 9525;

/** The fake measurer with external leading: a line of 16 px text is 17.6 px high. */
const gapMeasurer: TextMeasurer = {
  width: fakeMeasurer.width,
  metrics(font: LayoutFont): FontMetrics {
    return { ascent: font.size * 0.8, descent: font.size * 0.2, lineGap: font.size * 0.1 };
  },
};

const fld = (type: string) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
const instr = (code: string) => `<w:r><w:instrText xml:space="preserve">${code}</w:instrText></w:r>`;

const PICTURE =
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="x"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm></pic:spPr></pic:pic></a:graphicData></a:graphic>';
const RELS: [string, string, string][] = [['rIdImg', 'image', 'media/image1.png']];
const inline = (width: number, height: number, graphic = PICTURE) =>
  `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${width * PX}" cy="${height * PX}"/>` +
  `<wp:docPr id="1" name="Picture 1"/>${graphic}</wp:inline></w:drawing></w:r>`;
const images = (page: LayoutPage) => (page.items as ImageItem[]).filter((item) => item.type === 'image');

// ------------------------------------------------------------------ C1 fields

test('doc layout fix: a field in the code of another field shows nothing of its own', async () => {
  // A{ = { PAGE } + 10 }B with the cached result 17.
  const nested = fld('begin') + instr(' = ') + fld('begin') + instr(' PAGE ') + fld('separate') + r('1') + fld('end') + instr(' + 10 ') + fld('separate') + r('17') + fld('end');
  const { pages } = await lay(p(r('A') + nested + r('B')));
  assert.equal(
    texts(pages[0])
      .map((item) => item.text)
      .join(''),
    'A17B',
  );

  // In a footer: [{ IF { NUMPAGES } > 1 "many" "one" }] with the cached result "many".
  const condition =
    fld('begin') + instr(' IF ') + fld('begin') + instr(' NUMPAGES ') + fld('separate') + r('1') + fld('end') + instr(' > 1 "many" "one" ') + fld('separate') + r('many') + fld('end');
  const footer = await lay({
    body: p('text') + sectPr('<w:footerReference w:type="default" r:id="f1"/>'),
    footers: { f1: p(r('[') + condition + r(']')) },
  });
  const shown = texts(footer.pages[0])
    .filter((item) => item.source && item.source.story.indexOf('footer:') === 0)
    .map((item) => item.text)
    .join('');
  assert.equal(shown, '[many]');
  assert.equal(
    texts(footer.pages[0])
      .map((item) => item.text)
      .join(''),
    'text[many]',
  );
});

test('doc layout fix: a live field in the result of another field is still computed', async () => {
  // { IF 1 = 1 "{ PAGE }" } : the result holds a PAGE field of its own.
  const outer = fld('begin') + instr(' IF 1 = 1 ') + fld('separate') + r('p') + fld('begin') + instr(' PAGE ') + fld('separate') + r('9') + fld('end') + fld('end');
  const { pages } = await lay(p('first') + p(outer, '<w:pageBreakBefore/>'));
  assert.equal(
    texts(pages[1])
      .map((item) => item.text)
      .join(''),
    'p2',
  );
});

// ----------------------------------------------------------- C5 decimal tabs

const DECIMAL = '<w:tabs><w:tab w:val="decimal" w:pos="3000"/></w:tabs>';
const number = (text: string, lang?: string) => p(`<w:r>${lang ? `<w:rPr><w:lang w:val="${lang}"/></w:rPr>` : ''}<w:tab/><w:t>${text}</w:t></w:r>`, DECIMAL);
/** x of the character `char` of the only text item on each line of the page. */
const charX = (page: LayoutPage, char: string) => texts(page).map((item) => round(item.x + item.text.indexOf(char) * 8));

test('doc layout fix: a decimal tab aligns on the decimal separator of the language of the text', async () => {
  // Turkish: the comma is the decimal separator, the stop is 200 px into the column.
  const turkish = await lay(number('12.500,75', 'tr-TR') + number('500,75', 'tr-TR') + number('1.250.000,5', 'tr-TR'));
  assert.deepEqual(charX(turkish.pages[0], ','), [240, 240, 240]);
  assert.deepEqual(
    texts(turkish.pages[0]).map((item) => item.x),
    [240 - 6 * 8, 240 - 3 * 8, 240 - 9 * 8],
  );

  // English: the point; the commas group thousands.
  const english = await lay(number('1,250.5', 'en-US') + number('12,500.75', 'en-GB') + number('0.5', 'en-US'));
  assert.deepEqual(charX(english.pages[0], '.'), [240, 240, 240]);

  // A number without the separator of its language ends at the stop.
  const whole = await lay(number('1.250', 'tr-TR') + number('1,250', 'en-US'));
  assert.deepEqual(
    texts(whole.pages[0]).map((item) => item.x + item.width),
    [240, 240],
  );
});

test('doc layout fix: without a language on the run the default language decides, then the last separator of the number', async () => {
  const german = STYLES.replace('<w:sz w:val="24"/>', '<w:sz w:val="24"/><w:lang w:val="de-DE"/>');
  const styled = await lay({ styles: german, body: number('12.500,75') + number('1.250') });
  assert.equal(charX(styled.pages[0], ',')[0], 240);
  assert.equal(texts(styled.pages[0])[1].x + texts(styled.pages[0])[1].width, 240);

  const themed = await lay({ settings: '<w:themeFontLang w:val="tr-TR"/>', body: number('12.500,75') });
  assert.equal(charX(themed.pages[0], ',')[0], 240);

  // Nothing says what language: the last "." or "," between digits is taken for the decimal separator.
  const unknown = await lay(number('12.500,75') + number('1,250.5') + number('7,5') + number('1250'));
  assert.equal(charX(unknown.pages[0], ',')[0], 240);
  assert.equal(charX(unknown.pages[0], '.')[1], 240);
  assert.equal(charX(unknown.pages[0], ',')[2], 240);
  assert.equal(texts(unknown.pages[0])[3].x + 32, 240);
});

// ------------------------------------------------------ C7 hidden paragraphs

test('doc layout fix: a paragraph whose runs and mark are hidden takes no room', async () => {
  const hidden = '<w:p><w:pPr><w:spacing w:before="300" w:after="300"/><w:rPr><w:vanish/></w:rPr></w:pPr><w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r></w:p>';
  const empty = '<w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr></w:p>';
  const { pages } = await lay(p('one') + hidden + empty + p('two'));
  assert.deepEqual(at(pages[0]), ['one@40,52.8', 'two@40,68.8']);

  // A hidden mark alone does not hide text that shows; hidden text alone leaves the line of the mark.
  const mixed = await lay(p('one') + '<w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr><w:r><w:t>shown</w:t></w:r></w:p>' + p(r('gone', '<w:vanish/>')) + p('two'));
  assert.deepEqual(at(mixed.pages[0]), ['one@40,52.8', 'shown@40,68.8', 'two@40,100.8']);

  // In a cell, and as the only content of a document.
  const cell = await lay(tbl(tr(tc(hidden + p('in cell'))), [3000]));
  assert.deepEqual(at(cell.pages[0]), ['in cell@40,52.8']);
  const only = await lay(hidden);
  assert.equal(only.pages.length, 1);
  assert.equal(texts(only.pages[0]).length, 0);
});

test('doc layout fix: a hidden paragraph that ends a section still ends it', async () => {
  const mark = `<w:p><w:pPr><w:rPr><w:vanish/></w:rPr>${sectPr()}</w:pPr></w:p>`;
  const { pages } = await lay(p('one') + mark + p('two') + sectPr());
  assert.equal(pages.length, 2);
  assert.deepEqual(at(pages[1]), ['two@40,52.8']);
});

// -------------------------------------------------- B3 lines with a picture

test('doc layout fix: a line that holds only a picture is as high as the picture, without descent or leading', async () => {
  const { pages } = await lay({ body: p('before') + p(inline(100, 50)) + p('after'), rels: RELS }, undefined, gapMeasurer);
  // Lines of text are 17.6 px; the picture line is 50 px: no 3.2 px of descent and 1.6 px of leading under it.
  assert.deepEqual(at(pages[0]).map((entry) => entry.replace(/(\.\d)\d+$/, '$1')), ['before@40,52.8', 'after@40,120.4']);
  assert.deepEqual([images(pages[0])[0].x, round(images(pages[0])[0].y)], [40, 57.6]);

  // A picture lower than the text leaves the line as high as the text; it stands on the baseline.
  const small = await lay({ body: p('before') + p(inline(20, 10)) + p('after'), rels: RELS }, undefined, gapMeasurer);
  assert.deepEqual(baselines(small.pages[0]), [52.8, 88]);
  assert.equal(round(images(small.pages[0])[0].y), 60.4);

  // Between the two: the picture stands on a baseline that is as far down as it needs.
  const middle = await lay({ body: p('before') + p(inline(20, 15)) + p('after'), rels: RELS }, undefined, gapMeasurer);
  assert.deepEqual(baselines(middle.pages[0]), [52.8, 88]);
  assert.equal(round(images(middle.pages[0])[0].y), 57.6);
});

test('doc layout fix: text beside a picture keeps its descent and leading under the baseline', async () => {
  const { pages } = await lay({ body: p(r('ab') + inline(100, 50) + r('cd')) + p('next'), rels: RELS }, undefined, gapMeasurer);
  // 50 px above the baseline, 3.2 + 1.6 px below it.
  assert.deepEqual(at(pages[0]), ['ab@40,90', 'cd@156,90', 'next@40,107.6']);
  assert.deepEqual([images(pages[0])[0].x, images(pages[0])[0].y], [56, 40]);
});

test('doc layout fix: picture lines not at the end of the paragraph, and with line spacing', async () => {
  // Three pictures that each need a line of their own: 20 px a line, the last one too (the mark is lower).
  const stacked = await lay({ body: p(inline(300, 20) + inline(300, 20) + inline(300, 20)) + p('after'), rels: RELS }, undefined, gapMeasurer);
  assert.deepEqual(
    images(stacked.pages[0]).map((image) => round(image.y)),
    [40, 60, 80],
  );
  assert.deepEqual(baselines(stacked.pages[0]), [112.8]);

  // "At least" and exact spacing still rule the line.
  const atLeast = await lay({ body: p(inline(100, 50), '<w:spacing w:line="900" w:lineRule="atLeast"/>') + p('after'), rels: RELS }, undefined, gapMeasurer);
  assert.equal(round(images(atLeast.pages[0])[0].y), 50);
  assert.deepEqual(baselines(atLeast.pages[0]), [112.8]);
});

// ------------------------------------------------------- A1 very long words

/** Counts what measuring costs: the characters handed to width(). */
function countingMeasurer(): TextMeasurer & { measured: number; calls: number; longest: number } {
  const counter = {
    measured: 0,
    calls: 0,
    longest: 0,
    width(font: LayoutFont, text: string): number {
      counter.measured += text.length;
      counter.calls++;
      if (text.length > counter.longest) counter.longest = text.length;
      return fakeMeasurer.width(font, text);
    },
    metrics: fakeMeasurer.metrics,
  };
  return counter;
}

for (const length of [40000, 400000]) {
  test(`doc layout fix: a single word of ${length} characters is measured in linear time`, async () => {
    const measurer = countingMeasurer();
    const word = 'x'.repeat(length);
    const started = Date.now();
    const { pages } = await lay(p(word), undefined, measurer);
    const elapsed = Date.now() - started;
    // 40 characters a line, 20 lines a page.
    assert.equal(pages.length, length / 800);
    let shown = 0;
    for (const page of pages) for (const item of texts(page)) shown += item.text.length;
    assert.equal(shown, length);
    assert.equal(texts(pages[pages.length - 1]).pop()!.source!.offset, length - 40);
    // Once for the whole word, once more for each line it is cut into, and a little.
    assert.ok(measurer.measured <= 4 * length, `measured ${measurer.measured} characters for a word of ${length}`);
    assert.ok(elapsed < 20000, `took ${elapsed} ms`);
  });
}

test('doc layout fix: a long word after other text, with spaces after it and in a narrow cell', async () => {
  const measurer = countingMeasurer();
  const word = 'y'.repeat(5000);
  const { pages } = await lay(p(`ab ${word}  cd`), undefined, measurer);
  const all = pages.reduce<TextItem[]>((out, page) => out.concat(texts(page)), []);
  assert.equal(all.map((item) => item.text).join(''), `ab ${word}  cd`);
  // The word starts on a line of its own and fills every line but the last.
  assert.equal(all[0].text, 'ab ');
  assert.equal(all[1].text.length, 40);
  assert.ok(measurer.measured <= 4 * 5010, `measured ${measurer.measured}`);

  const cell = countingMeasurer();
  const table = await lay(tbl(tr(tc(word) + tc('z')), [1500, 1500]), undefined, cell);
  let shown = 0;
  for (const page of table.pages) for (const item of texts(page)) if (item.text.charAt(0) === 'y') shown += item.text.length;
  assert.equal(shown, 5000);
  assert.ok(cell.measured <= 6 * 5000, `measured ${cell.measured}`);
});

test('doc layout fix: a paragraph breaks into at most MAX_PARAGRAPH_LINES lines; the last one takes the rest', async () => {
  const count = MAX_PARAGRAPH_LINES + 5000;
  // One letter and a space per word, in a column one letter wide: every word is a line.
  const doc = await readSpec({ styles: STYLES, body: p('a '.repeat(count)) + sectPr() });
  const ctx = new LayoutContext(doc, fakeMeasurer, {});
  ctx.layoutBlocks = layoutBlocks;
  const story = { id: 'body', main: true };
  const prep = getPrep(ctx, doc.body.blocks[0] as Paragraph, story, [0], undefined, 0, undefined);
  const bands = paragraphBands(ctx, prep, 8, paragraphEnv(ctx, doc.body.blocks, 0, prep, undefined, undefined));
  assert.equal(bands.length, MAX_PARAGRAPH_LINES);
  let shown = 0;
  for (const band of bands) for (const item of band.items) if (item.type === 'text') shown += item.text.length;
  // Nothing is lost: the spaces at the end of a line are not shown, everything else is.
  assert.ok(shown >= 2 * count - MAX_PARAGRAPH_LINES - 1, `${shown} characters shown of ${2 * count}`);
  const last = bands[bands.length - 1].items.filter((item) => item.type === 'text') as TextItem[];
  assert.ok(last.reduce((sum, item) => sum + item.text.length, 0) > 5000);
});


// ----------------------------------------------------------- B1 table indent

const COMPAT14 = '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>';
const COMPAT15 = '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>';
const GRID_BORDERS =
  '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>`).join('') + '</w:tblBorders>';
const boxGraphic = (inner: string, spPr = '<wps:spPr/>', bodyPr = '<wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/>') =>
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>' +
  `${spPr}<wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx>${bodyPr}</wps:wsp></a:graphicData></a:graphic>`;
/** An anchored drawing `x` px into the column and `y` px below its paragraph. */
const anchored = (x: number, y: number, width: number, height: number, graphic: string, wrap = '<wp:wrapNone/>') =>
  '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>${x * PX}</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${y * PX}</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="${width * PX}" cy="${height * PX}"/>${wrap}<wp:docPr id="2" name="Anchored"/>${graphic}</wp:anchor></w:drawing></w:r>`;
/** Smallest and largest x of the vertical lines of a page. */
const verticalEdges = (page: LayoutPage) => {
  const xs = lines(page)
    .filter((line) => line.x1 === line.x2)
    .map((line) => round(line.x1));
  return [Math.min(...xs), Math.max(...xs)];
};

test('doc layout fix: before Word 2013 the edge of a table is its indent less the left padding of its first cell', async () => {
  // A text box from x = 50 to 350 holding a table of 290 px, indented by 100 px, whose first cell pads 100 px on the left.
  const table = tbl(tr(tc('cell', '<w:tcMar><w:left w:w="1500" w:type="dxa"/></w:tcMar>')), [4350], `<w:tblW w:w="4350" w:type="dxa"/><w:tblInd w:w="1500" w:type="dxa"/>${GRID_BORDERS}`);
  const body = p(anchored(10, 0, 300, 100, boxGraphic(table + p('')))) + p('');
  const old = await lay({ body, settings: COMPAT14 });
  assert.deepEqual(verticalEdges(old.pages[0]), [50, 340]);
  assert.equal(find(old.pages[0], 'cell').x, 150);

  // From Word 2013 on the indent is where the edge is.
  const modern = await lay({ body, settings: COMPAT15 });
  assert.deepEqual(verticalEdges(modern.pages[0]), [150, 440]);
  assert.equal(find(modern.pages[0], 'cell').x, 250);
});

test('doc layout fix: the padding that moves the table edge is that of the first cell of the first row', async () => {
  // Table default 10 px; the first cell of the first row states 20 px, the cells below state nothing.
  const rows = tr(tc('a1', '<w:tcMar><w:left w:w="300" w:type="dxa"/></w:tcMar>') + tc('b1')) + tr(tc('a2') + tc('b2'));
  const body = tbl(rows, [1500, 1500], `${CELL_MARGINS}${GRID_BORDERS}`);
  const old = await lay({ body, settings: COMPAT14 });
  // The text of the first cell lines up with the margin; the edge is 20 px left of it.
  assert.deepEqual(verticalEdges(old.pages[0]), [20, 220]);
  assert.deepEqual(
    texts(old.pages[0]).map((item) => `${item.text}@${item.x}`),
    ['a1@40', 'b1@130', 'a2@30', 'b2@130'],
  );

  // A table in a cell follows the same rule.
  const inner = tbl(tr(tc('in')), [900], `${CELL_MARGINS}<w:tblInd w:w="300" w:type="dxa"/>`);
  const nested = await lay({ body: tbl(tr(tc(inner + p('x'))), [3000]), settings: COMPAT14 });
  // Outer edge 30, outer padding 10: the cell text starts at 40; the inner table starts 20 - 10 px into it.
  assert.equal(find(nested.pages[0], 'in').x, 40 + 20);
  const modern = await lay({ body: tbl(tr(tc(inner + p('x'))), [3000]), settings: COMPAT15 });
  assert.equal(find(modern.pages[0], 'in').x, 50 + 20 + 10);
});

// ---------------------------------------------------------------- C tables

const modern = (body: string, measurer?: TextMeasurer) => lay({ body, settings: COMPAT15 }, undefined, measurer);
const rowHeight = (rule: string, twips: number) => `<w:trHeight w:val="${twips}" w:hRule="${rule}"/>`;

test('doc layout fix: a row of exact height shows only the lines that fit in it', async () => {
  // 40 px: two lines of 16 px fit, the third would run into the row below.
  const rows = tr(tc(p('e1') + p('e2') + p('e3') + p('e4')) + tc('x'), rowHeight('exact', 600)) + tr(tc('below') + tc('y'));
  const { pages } = await modern(tbl(rows, [1500, 1500]));
  assert.deepEqual(at(pages[0]), ['e1@50,52.8', 'e2@50,68.8', 'x@150,52.8', 'below@50,92.8', 'y@150,92.8']);

  // The first line is shown even when it is taller than the row: better cut than gone.
  const low = await modern(tbl(tr(tc(p('f1') + p('f2')), rowHeight('exact', 150)) + tr(tc('below')), [1500]));
  assert.deepEqual(at(low.pages[0]), ['f1@50,52.8', 'below@50,62.8']);
});

test('doc layout fix: a vertically merged cell taller than a page goes on from page to page', async () => {
  // 40 rows; the first column is one merged cell with 400 words (two to a line: 200 lines of 16 px).
  let rows = '';
  for (let i = 0; i < 40; i++) rows += tr(tc(i === 0 ? p(words(400)) : p(''), i === 0 ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>') + tc(`r${i}`));
  const { doc, pages } = await modern(tbl(rows, [1500, 1500]) + p('after'));
  assert.ok(pages.length >= 10, `${pages.length} pages`);
  let merged = '';
  for (const page of pages) {
    for (const item of texts(page)) {
      // Everything is in the text area of its page (40 to 360 px).
      assert.ok(item.y - 12.8 >= 40 - 0.01 && item.y + 3.2 <= 360 + 0.01, `text at y = ${item.y}`);
      if (item.source && item.source.path.length === 4 && item.source.path[2] === 0) merged += item.text;
    }
  }
  assert.equal(merged, words(400));
  // The rows of the second column are one line each, in order, and the paragraph after the table follows.
  const second = pages.reduce<string[]>((out, page) => out.concat(texts(page).filter((item) => item.x === 150).map((item) => item.text)), []);
  assert.deepEqual(second, Array.from({ length: 40 }, (_, i) => `r${i}`));
  assert.equal(texts(pages[pages.length - 1]).pop()!.text, 'after');
  assert.equal(doc.body.blocks.length, 2);
});

test('doc layout fix: a merged cell that fits is laid out as before; rows it spans may part between pages', async () => {
  const merged = (vAlign = '') =>
    tbl(tr(tc(p('m1') + p('m2'), '<w:vMerge w:val="restart"/>' + vAlign) + tc('r1')) + tr(tc(p(''), '<w:vMerge/>') + tc('r2')) + tr(tc(p(''), '<w:vMerge/>') + tc('r3')) + tr(tc('n') + tc('r4')), [1500, 1500]);
  const top = await modern(merged());
  assert.deepEqual(at(top.pages[0]).sort(), ['m1@50,52.8', 'm2@50,68.8', 'n@50,100.8', 'r1@150,52.8', 'r2@150,68.8', 'r3@150,84.8', 'r4@150,100.8']);
  const centre = await modern(merged('<w:vAlign w:val="center"/>'));
  assert.equal(find(centre.pages[0], 'm1').y, 52.8 + 8);
  assert.equal(find(centre.pages[0], 'm2').y, 68.8 + 8);
  const bottom = await modern(merged('<w:vAlign w:val="bottom"/>'));
  assert.equal(find(bottom.pages[0], 'm2').y, 84.8);

  // 19 lines above: the first row of the merged cell is the last line of the page, the rest follows on the next.
  const parted = await modern(Array.from({ length: 19 }, (_, i) => p(`L${i}`)).join('') + merged());
  assert.equal(pageText(parted.pages[0]).slice(-7), 'L18m1r1');
  assert.deepEqual(at(parted.pages[1]).sort(), ['m2@50,52.8', 'n@50,84.8', 'r2@150,52.8', 'r3@150,68.8', 'r4@150,84.8']);
});

test('doc layout fix: a line of a merged cell that straddles two rows is not cut by the page end', async () => {
  // Rows of 24 px, lines of 16 px: the second line of the merged cell starts in the first row and ends in the second.
  const rows =
    tr(tc(p('m1') + p('m2') + p('m3'), '<w:vMerge w:val="restart"/>') + tc('r1'), rowHeight('atLeast', 360)) +
    tr(tc(p(''), '<w:vMerge/>') + tc('r2'), rowHeight('atLeast', 360));
  // The table starts at y = 312 and ends at 360, the foot of the text area: it fits.
  const fits = await modern(p('x', '<w:spacing w:before="3840"/>') + tbl(rows, [1500, 1500]));
  assert.equal(fits.pages.length, 1);
  assert.deepEqual(at(fits.pages[0]).slice(1).sort(), ['m1@50,324.8', 'm2@50,340.8', 'm3@50,356.8', 'r1@150,324.8', 'r2@150,348.8']);
  // 16 px lower the first row still fits (it ends at 352) and the second goes to the next page.
  const parts = await modern(p('x', '<w:spacing w:before="4080"/>') + tbl(rows, [1500, 1500]));
  assert.equal(pageText(parts.pages[0]), 'xm1m2r1');
  assert.deepEqual(at(parts.pages[1]), ['m3@50,60.8', 'r2@150,52.8']);
  // 8 px more: the first row would end at 360 with the second line of the merged cell hanging out of the text area.
  const hangs = await modern(p('x', '<w:spacing w:before="4200"/>') + tbl(rows, [1500, 1500]));
  for (const page of hangs.pages) for (const item of texts(page)) assert.ok(item.y + 3.2 <= 360.01, `"${item.text}" at y = ${item.y}`);
  assert.equal(
    hangs.pages
      .map((page) => texts(page).filter((item) => item.text.charAt(0) === 'm').map((item) => item.text).join(''))
      .join('|'),
    'm1|m2m3',
  );
});

test('doc layout fix: text in a cell with vertical text direction is laid out along the height of the row and turned', async () => {
  const cell = (direction: string, extra = '') => tc('vertical', `<w:textDirection w:val="${direction}"/>${extra}`);
  // Bottom to top: the line starts at the foot of the row; lines stack from the left.
  const up = await modern(tbl(tr(cell('btLr') + tc('x'), rowHeight('atLeast', 1500)) + tr(tc('below') + tc('y')), [1500, 1500]));
  const upText = find(up.pages[0], 'vertical');
  assert.deepEqual([upText.rotation, round(upText.x), upText.y, upText.width], [270, 62.8, 140, 64]);
  assert.equal(find(up.pages[0], 'below').y, 152.8);
  assert.deepEqual(upText.source, { story: 'body', path: [0, 0, 0, 0], offset: 0 });

  // Top to bottom: the line starts at the top; lines stack from the right.
  const down = await modern(tbl(tr(cell('tbRl') + tc('x'), rowHeight('atLeast', 1500)), [1500, 1500]));
  const downText = find(down.pages[0], 'vertical');
  assert.deepEqual([downText.rotation, round(downText.x), downText.y, downText.width], [90, 117.2, 40, 64]);

  // Without a height of its own the row is as high as the text is long; the text is not cut letter by letter.
  const free = await modern(tbl(tr(cell('btLr') + tc('x')) + tr(tc('b') + tc('y')), [600, 1500]));
  const all = texts(free.pages[0]).filter((item) => item.rotation !== undefined);
  assert.deepEqual(all.map((item) => item.text), ['vertical']);
  assert.equal(all[0].y, 40 + 64);
  assert.equal(find(free.pages[0], 'b').y, 40 + 64 + 12.8);

  // Vertical alignment moves the lines across the cell; horizontal text is untouched.
  const centred = await modern(tbl(tr(cell('btLr', '<w:vAlign w:val="center"/>') + tc('x'), rowHeight('atLeast', 1500)), [1500, 1500]));
  assert.equal(round(find(centred.pages[0], 'vertical').x), 50 + (80 - 16) / 2 + 12.8);
  assert.equal(find(centred.pages[0], 'x').rotation, undefined);
});

test('doc layout fix: shading and borders of a paragraph in a vertical cell turn with the text', async () => {
  const shaded = p('ab', '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00"/>');
  const { pages } = await modern(tbl(tr(tc(shaded, '<w:textDirection w:val="tbRl"/>'), rowHeight('exact', 1500)), [1500]));
  const [rect] = rects(pages[0]);
  // The line is 16 px thick and 100 px long: a strip at the right edge of the cell content, down the whole row.
  assert.deepEqual([rect.x, rect.y, rect.width, rect.height], [50 + 80 - 16, 40, 16, 100]);
});

test('doc layout fix: a long table is laid out a few rows at a time', async () => {
  const measurer = countingMeasurer();
  let rows = '';
  for (let i = 0; i < 3000; i++) rows += tr(tc(`a${i}`) + tc(`b${i}`));
  const doc = await readSpec({ styles: STYLES, settings: COMPAT15, body: tbl(rows, [1500, 1500]) + p('after') + sectPr() });
  const layout = new DocLayout(doc, measurer, {});
  assert.equal(layout.layout({ pages: 1 }).done, false);
  // The first page holds 20 rows; far fewer cells than the table has were measured for it.
  assert.equal(texts(layout.pages[0]).length, 40);
  assert.ok(measurer.calls < 1000, `${measurer.calls} texts measured for the first page`);
  assert.equal(layout.layout().done, true);
  assert.equal(layout.pages.length, 151);
  assert.deepEqual(
    texts(layout.pages[150]).map((item) => item.text),
    ['after'],
  );
  assert.equal(texts(layout.pages[149]).pop()!.text, 'b2999');
});

// ------------------------------------------------------- C4 sides of a wrap

const square = (wrapText: string | undefined, x: number) =>
  anchored(x, 0, 100, 50, PICTURE, wrapText === undefined ? '<wp:wrapSquare/>' : `<wp:wrapSquare wrapText="${wrapText}"/>`);
/** [x, number of characters] of the text items of the first lines of a page. */
const runs = (page: LayoutPage, count: number) =>
  texts(page)
    .slice(0, count)
    .map((item) => [item.x, round(item.y), item.text.replace(/ +$/, '').length]);

test('doc layout fix: text wraps on both sides of a drawing in the middle of the column', async () => {
  // The drawing takes x = 150..250 of a column from 40 to 360: 110 px (two words) are free on each side.
  const { pages } = await lay({ body: p(square('bothSides', 110) + r(words(40))), rels: RELS });
  assert.deepEqual(runs(pages[0], 10), [
    [40, 52.8, 9],
    [250, 52.8, 9],
    [40, 68.8, 9],
    [250, 68.8, 9],
    [40, 84.8, 9],
    [250, 84.8, 9],
    [40, 100.8, 9],
    [250, 100.8, 9],
    // Below the drawing: the whole width.
    [40, 116.8, 39],
    [40, 132.8, 39],
  ]);
  assert.equal(pageText(pages[0]), words(40));
  // No wrapText is both sides.
  const plain = await lay({ body: p(square(undefined, 110) + r(words(40))), rels: RELS });
  assert.deepEqual(runs(plain.pages[0], 10), runs(pages[0], 10));
});

test('doc layout fix: wrapText left, right and largest keep the text to one side', async () => {
  const left = await lay({ body: p(square('left', 110) + r(words(40))), rels: RELS });
  assert.deepEqual(runs(left.pages[0], 5), [
    [40, 52.8, 9],
    [40, 68.8, 9],
    [40, 84.8, 9],
    [40, 100.8, 9],
    [40, 116.8, 39],
  ]);
  const right = await lay({ body: p(square('right', 110) + r(words(40))), rels: RELS });
  assert.deepEqual(runs(right.pages[0], 5), [
    [250, 52.8, 9],
    [250, 68.8, 9],
    [250, 84.8, 9],
    [250, 100.8, 9],
    [40, 116.8, 39],
  ]);
  // 100 px on the left, 120 on the right: the right side is the larger one.
  const largest = await lay({ body: p(square('largest', 100) + r(words(40))), rels: RELS });
  assert.deepEqual(runs(largest.pages[0], 5), [
    [240, 52.8, 14],
    [240, 68.8, 14],
    [240, 84.8, 14],
    [240, 100.8, 14],
    [40, 116.8, 39],
  ]);
  const both = await lay({ body: p(square('bothSides', 100) + r(words(40))), rels: RELS });
  assert.deepEqual(runs(both.pages[0], 3), [
    [40, 52.8, 9],
    [240, 52.8, 14],
    [40, 68.8, 9],
  ]);
  for (const laid of [left, right, largest, both]) assert.equal(pageText(laid.pages[0]), words(40));
});

test('doc layout fix: beside a drawing a word too long for one side goes to the other; alignment holds on each side', async () => {
  // Thirty letters fit on neither side (13 letters each). The word does not follow "ab" on its line; on the next
  // line it is not cut into the left side but forced into the last one, the right; what is left of it fits on the left.
  const long = await lay({ body: p(square('bothSides', 110) + r('ab ' + 'x'.repeat(30) + ' cd')), rels: RELS });
  assert.deepEqual(runs(long.pages[0], 4), [
    [40, 52.8, 2],
    [250, 68.8, 13],
    [250, 84.8, 13],
    [40, 100.8, 7],
  ]);
  assert.equal(pageText(long.pages[0]), 'ab ' + 'x'.repeat(30) + ' cd');
  // Right aligned: each part ends at the right edge of its side.
  const aligned = await lay({ body: p(square('bothSides', 110) + r(words(8)), '<w:jc w:val="right"/>'), rels: RELS });
  assert.deepEqual(
    texts(aligned.pages[0])
      .slice(0, 2)
      .map((item) => [item.x, item.text]),
    [
      [150 - 72, 'www0 www1 '],
      [360 - 72, 'www2 www3 '],
    ],
  );
  // A numbered paragraph has its number once.
  const numbering =
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1800" w:hanging="300"/></w:pPr></w:lvl></w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
  const listed = await lay({ numbering, body: p(square('bothSides', 0) + r(words(6)), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'), rels: RELS });
  assert.equal(texts(listed.pages[0]).filter((item) => item.text === '1.').length, 1);
});

// ------------------------------------------------ C8 fill and frame of boxes

const frame = (page: LayoutPage) =>
  lines(page)
    .map((line) => [round(line.x1), round(line.y1), round(line.x2), round(line.y2), line.color, round(line.width)].join(' '))
    .sort();

test('doc layout fix: a text box shows the fill and the outline of its shape', async () => {
  const spPr =
    '<wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFEEDD"/></a:solidFill>' +
    '<a:ln w="19050"><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:ln></wps:spPr>';
  const { pages } = await lay({ body: p(anchored(10, 0, 100, 50, boxGraphic(p('in'), spPr))) });
  const page = pages[0];
  const [fill] = rects(page);
  assert.deepEqual([fill.x, fill.y, fill.width, fill.height, fill.fill], [50, 40, 100, 50, '#FFEEDD']);
  assert.deepEqual(frame(page), ['150 40 150 90 #112233 2', '50 40 150 40 #112233 2', '50 40 50 90 #112233 2', '50 90 150 90 #112233 2']);
  // The fill lies under the text of the box, the outline over it.
  const text = find(page, 'in');
  assert.ok(page.items.indexOf(fill) < page.items.indexOf(text));
  assert.ok(page.items.indexOf(lines(page)[0]) > page.items.indexOf(text));
});

test('doc layout fix: theme colours, missing parts and "no fill" of a text box', async () => {
  // A theme colour made lighter; an outline without a width is 0.75 pt.
  const themed =
    '<wps:spPr><a:solidFill><a:schemeClr val="bg1"><a:lumMod val="85000"/></a:schemeClr></a:solidFill><a:ln><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln></wps:spPr>';
  const first = await lay({ theme: THEME, body: p(anchored(10, 0, 100, 50, boxGraphic(p('in'), themed))) });
  assert.equal(rects(first.pages[0])[0].fill, '#D9D9D9');
  assert.deepEqual(frame(first.pages[0])[0], '150 40 150 90 #4472C4 1');

  // Nothing stated, or stated as nothing: no fill and no outline.
  for (const spPr of ['<wps:spPr/>', '<wps:spPr><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>', '<wps:spPr><a:solidFill><a:srgbClr val="ZZZZZZ"/></a:solidFill></wps:spPr>']) {
    const none = await lay({ body: p(anchored(10, 0, 100, 50, boxGraphic(p('in'), spPr))) });
    assert.equal(rects(none.pages[0]).length, 0, spPr);
    assert.equal(lines(none.pages[0]).length, 0, spPr);
  }

  // The style of the shape gives the colours its properties do not state.
  const styled =
    '<wps:spPr/><wps:style><a:lnRef idx="2"><a:srgbClr val="00AA00"/></a:lnRef><a:fillRef idx="1"><a:srgbClr val="EEEEEE"/></a:fillRef>' +
    '<a:effectRef idx="0"><a:srgbClr val="000000"/></a:effectRef><a:fontRef idx="minor"><a:srgbClr val="000000"/></a:fontRef></wps:style>';
  const byStyle = await lay({ body: p(anchored(10, 0, 100, 50, boxGraphic(p('in'), styled))) });
  assert.equal(rects(byStyle.pages[0])[0].fill, '#EEEEEE');
  assert.equal(lines(byStyle.pages[0])[0].color, '#00AA00');

  // A box without a height of its own is framed around its text.
  const auto = await lay({ body: p(anchored(10, 0, 100, 0, boxGraphic(p('a') + p('b'), '<wps:spPr><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></wps:spPr>'))) });
  assert.equal(rects(auto.pages[0])[0].height, 32);
});

test('doc layout fix: a VML text box shows its fill and stroke', async () => {
  const vml = (attrs: string) =>
    `<w:r><w:pict><v:shape style="width:75pt;height:37.5pt" ${attrs}><v:textbox><w:txbxContent>${p('in')}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`;
  const coloured = await lay({ body: p(vml('fillcolor="#ffeedd" strokecolor="#123 [3204]" strokeweight="1.5pt"')) });
  const [fill] = rects(coloured.pages[0]);
  assert.deepEqual([fill.x, fill.width, fill.height, fill.fill], [40, 100, 50, '#FFEEDD']);
  assert.equal(lines(coloured.pages[0]).length, 4);
  assert.deepEqual([lines(coloured.pages[0])[0].color, lines(coloured.pages[0])[0].width], ['#112233', 2]);

  // VML shapes are outlined in black unless they say otherwise; the white fill they have by default is not painted.
  const plain = await lay({ body: p(vml('')) });
  assert.equal(rects(plain.pages[0]).length, 0);
  assert.deepEqual([lines(plain.pages[0]).length, lines(plain.pages[0])[0].color, lines(plain.pages[0])[0].width], [4, '#000000', 1]);
  const bare = await lay({ body: p(vml('filled="f" stroked="f" fillcolor="red"')) });
  assert.equal(rects(bare.pages[0]).length + lines(bare.pages[0]).length, 0);
  const named = await lay({ body: p(vml('fillcolor="yellow" stroked="false"')) });
  assert.equal(rects(named.pages[0])[0].fill, '#FFFF00');
  assert.equal(lines(named.pages[0]).length, 0);
});
