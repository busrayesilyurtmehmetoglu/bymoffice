import assert from 'node:assert/strict';
import { test } from 'node:test';
import type { DocxSpec } from './docx-fixtures';
import { CELL_MARGINS, COMPAT15, STYLES, at, baselines, filler, find, lay, lines, p, pageText, rects, round, tbl, tc, texts, tr } from './doc-layout-fixtures';

// Tables in a document of Word 2013 start at the margin (x = 40); cells have 10 px of margin left and right.
const modern = (body: string, more: Partial<DocxSpec> = {}) => lay({ body, settings: COMPAT15, ...more });

test('doc layout: a table takes its column widths from the grid; rows are as high as their tallest cell', async () => {
  const table = tbl(tr(tc('aaaa bbbb cccc') + tc('one')) + tr(tc('x') + tc('y')), [1500, 3000]);
  const { pages } = await modern(table + p('after'));
  assert.deepEqual(at(pages[0]), [
    // 100 px less margins: 10 characters a line.
    'aaaa bbbb @50,52.8',
    'cccc@50,68.8',
    'one@150,52.8',
    'x@50,84.8',
    'y@150,84.8',
    'after@40,100.8',
  ]);
  assert.deepEqual(find(pages[0], 'one').source, { story: 'body', path: [0, 0, 1, 0], offset: 0 });

  // Documents of earlier versions line the text up with the margin instead of the edge of the table.
  const old = await lay(table);
  assert.equal(find(old.pages[0], 'aaaa').x, 40);
});

test('doc layout: cell widths win over the grid; table width, indent and alignment', async () => {
  const w = (twips: number) => `<w:tcW w:w="${twips}" w:type="dxa"/>`;
  const cells = await modern(tbl(tr(tc('a', w(1500)) + tc('b', w(3000))), [1000, 1000]));
  assert.equal(find(cells.pages[0], 'b').x, 150);

  const pct = await modern(tbl(tr(tc('a') + tc('b')), [1000, 1000], CELL_MARGINS + '<w:tblW w:w="5000" w:type="pct"/>'));
  // 100 % of 320 px: two columns of 160.
  assert.equal(find(pct.pages[0], 'b').x, 40 + 160 + 10);
  const dxa = await modern(tbl(tr(tc('a') + tc('b')), [1000, 3000], CELL_MARGINS + '<w:tblW w:w="3000" w:type="dxa"/>'));
  assert.equal(find(dxa.pages[0], 'b').x, 40 + 50 + 10);

  const indent = await modern(tbl(tr(tc('a')), [1500], CELL_MARGINS + '<w:tblInd w:w="300" w:type="dxa"/>'));
  assert.equal(find(indent.pages[0], 'a').x, 70);
  const centred = await modern(tbl(tr(tc('a')), [3000], CELL_MARGINS + '<w:jc w:val="center"/>'));
  assert.equal(find(centred.pages[0], 'a').x, 40 + 60 + 10);
  const right = await modern(tbl(tr(tc('a')), [3000], CELL_MARGINS + '<w:jc w:val="right"/>'));
  assert.equal(find(right.pages[0], 'a').x, 40 + 120 + 10);
});

test('doc layout: automatic layout sizes columns nothing else sizes by their content', async () => {
  const rows = tr(tc('aaaa') + tc('bbbbbbbbbbbb')) + tr(tc('cc') + tc('dd'));
  const natural = await modern(tbl(rows, []));
  // 32 and 96 px of text plus margins.
  assert.equal(find(natural.pages[0], 'bbbb').x, 40 + 52 + 10);
  assert.equal(find(natural.pages[0], 'dd').x, 40 + 52 + 10);
  const full = await modern(tbl(rows, [], CELL_MARGINS + '<w:tblW w:w="5000" w:type="pct"/>'));
  // Stretched to the width of the table in proportion.
  assert.equal(round(find(full.pages[0], 'bbbb').x), round(40 + (52 * 320) / 168 + 10));
  // A table wider than the page is scaled down to it.
  const wide = await modern(tbl(tr(tc('a') + tc('b')), [6000, 6000]));
  assert.equal(find(wide.pages[0], 'b').x, 40 + 160 + 10);
  // Fixed layout keeps its widths.
  const fixed = await modern(tbl(tr(tc('a') + tc('b')), [6000, 6000], CELL_MARGINS + '<w:tblLayout w:type="fixed"/>'));
  assert.equal(find(fixed.pages[0], 'b').x, 40 + 400 + 10);
});

test('doc layout: cells that span columns and rows', async () => {
  const span = await modern(tbl(tr(tc('wide wide wide wide wide wide wide wide', '<w:gridSpan w:val="2"/>')) + tr(tc('a') + tc('b')) + tr(tc('c'), '<w:gridBefore w:val="1"/>'), [1500, 3000]));
  // 300 px less margins: 35 characters.
  assert.deepEqual(at(span.pages[0]), ['wide wide wide wide wide wide wide @50,52.8', 'wide@50,68.8', 'a@50,84.8', 'b@150,84.8', 'c@150,100.8']);

  const merged = await modern(
    tbl(tr(tc(p('m1') + p('m2') + p('m3'), '<w:vMerge w:val="restart"/>') + tc('r1')) + tr(tc(p(''), '<w:vMerge/>') + tc('r2')), [1500, 1500]) + p('after'),
  );
  // The merged cell needs 48 px; the second row grows to make room. Each row carries the lines that start in it.
  assert.deepEqual(at(merged.pages[0]), ['m1@50,52.8', 'r1@150,52.8', 'm2@50,68.8', 'm3@50,84.8', 'r2@150,68.8', 'after@40,100.8']);
});

test('doc layout: row heights, at least and exact; vertical alignment', async () => {
  const height = (rule: string, twips: number) => `<w:trHeight w:val="${twips}" w:hRule="${rule}"/>`;
  const { pages } = await modern(
    tbl(
      tr(tc('top') + tc('mid', '<w:vAlign w:val="center"/>') + tc('low', '<w:vAlign w:val="bottom"/>'), height('atLeast', 600)) +
        tr(tc(p('e1') + p('e2')) + tc('') + tc(''), height('exact', 150)) +
        tr(tc('last') + tc('') + tc('')),
      [1000, 1000, 1000],
    ),
  );
  // 40 px for the first row, 10 for the second whatever is in it.
  assert.equal(find(pages[0], 'top').y, 52.8);
  assert.equal(find(pages[0], 'mid').y, 52.8 + 12);
  assert.equal(find(pages[0], 'low').y, 52.8 + 24);
  assert.equal(find(pages[0], 'e1').y, 92.8);
  assert.equal(find(pages[0], 'last').y, 102.8);
});

test('doc layout: table borders, the heavier border where two cells meet, shading', async () => {
  const all = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((side) => `<w:${side} w:val="single" w:sz="8" w:space="0" w:color="auto"/>`).join('');
  const heavy = '<w:tcBorders><w:bottom w:val="single" w:sz="24" w:space="0" w:color="FF0000"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="CCCCCC"/>';
  const { pages } = await modern(tbl(tr(tc('a', heavy) + tc('b')) + tr(tc('c') + tc('d')), [1500, 1500], CELL_MARGINS + `<w:tblBorders>${all}</w:tblBorders>`) + p('after'));
  const page = pages[0];
  const thin = 4 / 3;
  // Borders take room: the line above each row, and the one under the table.
  assert.deepEqual(baselines(page), [round(52.8 + thin), round(52.8 + thin + 16 + 4), round(52.8 + thin + 16 + 4 + 16 + thin)]);
  const drawn = lines(page);
  // Per row: a top and a left line for each cell and the right line of the last; under the table two bottom lines.
  assert.equal(drawn.length, 12);
  const red = drawn.filter((line) => line.color === '#FF0000');
  assert.deepEqual(
    red.map((line) => [line.x1, round(line.y1), line.x2, round(line.y2), line.width]),
    [[40, round(40 + thin + 16 + 2), 140, round(40 + thin + 16 + 2), 4]],
  );
  const [shade] = rects(page);
  assert.deepEqual([shade.fill, shade.x, shade.y, shade.width, round(shade.height)], ['#CCCCCC', 40, 40, 100, round(16 + thin)]);
  // Verticals: at 40, 140 and 240.
  const xs = drawn.filter((line) => line.x1 === line.x2).map((line) => line.x1);
  assert.deepEqual(xs, [40, 140, 240, 40, 140, 240]);
});

test('doc layout: space between cells gives every cell its own box', async () => {
  const inside = ['insideH', 'insideV'].map((side) => `<w:${side} w:val="single" w:sz="8" w:space="0" w:color="auto"/>`).join('');
  const { pages } = await modern(tbl(tr(tc('a') + tc('b')), [1500, 1500], CELL_MARGINS + `<w:tblCellSpacing w:w="150" w:type="dxa"/><w:tblBorders>${inside}</w:tblBorders>`) + p('after'));
  // 10 px around every cell, and the room of its own borders above and below.
  assert.deepEqual(at(pages[0]), ['a@60,64.13', 'b@160,64.13', 'after@40,91.47']);
  assert.equal(lines(pages[0]).length, 8);
});

test('doc layout: a row is divided between pages unless it must not be', async () => {
  const cell = tc(p('c1') + p('c2') + p('c3') + p('c4')) + tc('other');
  const split = await modern(filler(18) + tbl(tr(cell), [1500, 1500]));
  assert.equal(split.pages.length, 2);
  assert.equal(pageText(split.pages[0]).slice(-12), 'L17c1c2other');
  assert.deepEqual(at(split.pages[1]), ['c3@50,52.8', 'c4@50,68.8']);

  const whole = await modern(filler(18) + tbl(tr(cell, '<w:cantSplit/>'), [1500, 1500]));
  assert.equal(pageText(whole.pages[0]).slice(-3), 'L17');
  assert.equal(pageText(whole.pages[1]), 'c1c2c3c4other');

  // A row taller than a page is divided however often it takes.
  const tall = await modern(tbl(tr(tc(filler(45, 'T'))), [1500]));
  assert.deepEqual(tall.pages.map((page) => texts(page).length), [20, 20, 5]);
  // And when it must not be divided it is placed anyway, running off the page, and layout goes on.
  const stuck = await modern(tbl(tr(tc(filler(45, 'T')), '<w:cantSplit/>'), [1500]) + p('after'));
  assert.equal(stuck.pages.length, 2);
  assert.equal(texts(stuck.pages[0]).length, 45);
  assert.deepEqual(at(stuck.pages[1]), ['after@40,52.8']);
});

test('doc layout: a heading that keeps with next stays with the first row of the table after it; a tall merged cell is divided between pages', async () => {
  const { pages } = await modern(filler(19) + p('Heading', '<w:keepNext/>') + tbl(tr(tc(p('r1a') + p('r1b'))) + tr(tc('r2')), [1500]));
  // The heading fits on the last line of the page, the row after it does not: the heading goes with the row.
  assert.equal(pageText(pages[0]).slice(-3), 'L18');
  assert.equal(pageText(pages[1]), 'Headingr1ar1br2');

  const merged = tbl(tr(tc(p('m1') + p('m2') + p('m3'), '<w:vMerge w:val="restart"/>') + tc('r1')) + tr(tc(p(''), '<w:vMerge/>') + tc('r2')), [1500, 1500]);
  const tall = await modern(filler(18) + merged);
  // Two lines are left on the page: the first row and the first line of the second, which is divided as any row is.
  // The merged cell goes on with its third line on the next page, as it does in Word.
  assert.equal(pageText(tall.pages[0]).slice(-11), 'L17m1r1m2r2');
  assert.equal(pageText(tall.pages[1]), 'm3');
});

test('doc layout: header rows repeat at the top of every page of the table', async () => {
  let rows = tr(tc('Head'), '<w:tblHeader/>');
  for (let i = 0; i < 25; i++) rows += tr(tc('R' + i));
  const { pages } = await modern(tbl(rows, [1500]));
  assert.equal(pages.length, 2);
  assert.equal(texts(pages[0]).length, 20);
  assert.deepEqual(at(pages[1]).slice(0, 2), ['Head@50,52.8', 'R19@50,68.8']);
  assert.equal(texts(pages[1])[0].repeat, true);
  assert.equal(texts(pages[0])[0].repeat, undefined);
  // The repeated header is not text of its own: every row is there once.
  assert.equal(pages.map(pageText).join(''), 'Head' + Array.from({ length: 25 }, (_, i) => 'R' + i).join(''));
});

test('doc layout: the table is closed with a line where it goes on to the next page', async () => {
  const borders = ['top', 'bottom', 'insideH'].map((side) => `<w:${side} w:val="single" w:sz="8" w:space="0" w:color="auto"/>`).join('');
  let rows = '';
  for (let i = 0; i < 25; i++) rows += tr(tc('R' + i));
  const { pages } = await modern(tbl(rows, [1500], CELL_MARGINS + `<w:tblBorders>${borders}</w:tblBorders>`));
  assert.equal(pages.length, 2);
  const first = lines(pages[0]);
  // 18 rows of 17.33 px fit; each has a line above it, and one closes the page.
  assert.equal(texts(pages[0]).length, 18);
  assert.equal(first.length, 19);
  assert.equal(round(first[18].y1), round(40 + 18 * (16 + 4 / 3)));
});

test('doc layout: a table positioned on the page sits where it says, and the text keeps clear of it', async () => {
  const placed = CELL_MARGINS + '<w:tblpPr w:leftFromText="0" w:rightFromText="0" w:vertAnchor="page" w:horzAnchor="page" w:tblpX="1500" w:tblpY="3000"/>';
  const { pages } = await modern(p('one') + tbl(tr(tc('float')), [1500], placed) + filler(12));
  const page = pages[0];
  // 100 px from the left edge of the paper and 200 px from its top.
  assert.deepEqual([find(page, 'float').x, find(page, 'float').y], [110, 212.8]);
  assert.deepEqual(find(page, 'float').source!.path, [1, 0, 0, 0]);
  // The line of text at that height starts right of the table; the others are where they would be without it.
  assert.deepEqual([find(page, 'L8').x, find(page, 'L8').y], [40, 196.8]);
  assert.deepEqual([find(page, 'L9').x, find(page, 'L9').y], [200, 212.8]);
  assert.deepEqual([find(page, 'L10').x, find(page, 'L10').y], [40, 228.8]);

  // Positioned relative to the text, a floating table stays in the flow.
  const inFlow = CELL_MARGINS + '<w:tblpPr w:vertAnchor="text" w:horzAnchor="margin" w:tblpY="1"/>';
  const flow = await modern(p('one') + tbl(tr(tc('float')), [1500], inFlow) + p('after'));
  assert.deepEqual(at(flow.pages[0]), ['one@40,52.8', 'float@50,68.8', 'after@40,84.8']);
});

test('doc layout: tables in tables', async () => {
  const inner = tbl(tr(tc('in1') + tc('in2')), [1000, 1000]);
  const { pages } = await modern(tbl(tr(tc('out') + tc(inner + p('below'))), [1500, 3000]));
  assert.deepEqual(at(pages[0]), ['out@50,52.8', 'in1@160,52.8', 'in2@226.67,52.8', 'below@150,68.8']);
  assert.deepEqual(find(pages[0], 'in2').source!.path, [0, 0, 1, 0, 0, 1, 0]);
});

test('doc layout: the conditional formatting of a table style reaches the text', async () => {
  const styles =
    STYLES +
    '<w:style w:type="table" w:styleId="T"><w:name w:val="T"/><w:tblPr><w:tblCellMar><w:left w:w="150" w:type="dxa"/><w:right w:w="150" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
    '<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="000080"/></w:tcPr></w:tblStylePr>' +
    '<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/></w:tcPr></w:tblStylePr></w:style>';
  const look = '<w:tblStyle w:val="T"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>';
  const { pages } = await modern(tbl(tr(tc('h')) + tr(tc('a')) + tr(tc('b')) + tr(tc('c')), [1500], look), { styles });
  assert.deepEqual(
    texts(pages[0]).map((item) => item.font.bold),
    [true, false, false, false],
  );
  assert.deepEqual(
    rects(pages[0]).map((rect) => [rect.fill, rect.y]),
    [
      ['#000080', 40],
      ['#EEEEEE', 56],
      ['#EEEEEE', 88],
    ],
  );
});
