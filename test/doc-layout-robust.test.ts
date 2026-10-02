import assert from 'node:assert/strict';
import { test } from 'node:test';
import { DocLayout } from '../src/layout/doc/layout';
import type { LayoutPage, TextItem, TextMeasurer } from '../src/layout/doc/types';
import { paragraphText, type Block, type DocModel, type Paragraph } from '../src/model/doc/types';
import { readSpec, type DocxSpec } from './docx-fixtures';
import { CELL_MARGINS, COMPAT15, STYLES, fakeMeasurer, filler, lay, p, r, sectPr, tbl, tc, texts, tr, words } from './doc-layout-fixtures';

/** Small deterministic generator (mulberry32). */
function random(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const field = (code: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>${code}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r('8')}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

/** A random document of paragraphs and tables; `columns` allows sections with more than one column. */
function randomSpec(seed: number, columns: boolean): DocxSpec {
  const next = random(seed);
  const int = (max: number) => Math.floor(next() * max);
  const pick = <T>(list: T[]): T => list[int(list.length)];
  let serial = 0;
  const text = (count: number) => {
    const out: string[] = [];
    for (let i = 0; i < count; i++) {
      const length = next() < 0.03 ? 60 + int(60) : 1 + int(9);
      let word = 'w' + (serial++).toString(36);
      while (word.length < length) word += pick(['a', 'b', '-', 'c', 'd']);
      out.push(word);
    }
    return out.join(' ');
  };
  const runs = (count: number) => {
    let out = '';
    const pieces = 1 + int(3);
    for (let k = 0; k < pieces; k++) {
      const rPr = pick(['', '', '<w:b/>', '<w:sz w:val="36"/>', '<w:i/><w:u w:val="single"/>', '<w:vertAlign w:val="superscript"/>', '<w:sz w:val="16"/>']);
      out += r(text(1 + int(count)) + (next() < 0.7 ? ' ' : ''), rPr);
      if (next() < 0.08) out += '<w:r><w:tab/></w:r>';
      if (next() < 0.04) out += '<w:r><w:br/></w:r>';
      if (next() < 0.01) out += '<w:r><w:br w:type="page"/></w:r>';
    }
    return out;
  };
  const paragraph = (inCell: boolean, small = false) => {
    // Rows that cannot be divided hold little, so that nothing has to be forced onto a page.
    if (small) return p(text(2).slice(0, 12));
    let pPr = '';
    if (next() < 0.2) pPr += pick(['<w:keepNext/>', '<w:keepLines/>', '<w:widowControl/>', '<w:keepNext/><w:keepLines/>']);
    if (!inCell && next() < 0.03) pPr += '<w:pageBreakBefore/>';
    if (next() < 0.3) pPr += `<w:spacing w:before="${int(400)}" w:after="${int(400)}" w:line="${pick([240, 276, 360, 480])}" w:lineRule="auto"/>`;
    else if (next() < 0.1) pPr += `<w:spacing w:line="${200 + int(300)}" w:lineRule="${pick(['exact', 'atLeast'])}"/>`;
    if (next() < 0.3) pPr += `<w:ind w:left="${int(700)}" w:right="${int(500)}" ${next() < 0.5 ? `w:firstLine="${int(400)}"` : `w:hanging="${int(300)}"`}/>`;
    if (next() < 0.4) pPr += `<w:jc w:val="${pick(['left', 'center', 'right', 'both'])}"/>`;
    return p(next() < 0.1 ? '' : runs(inCell ? 5 : next() < 0.2 ? 60 : 12), pPr);
  };
  const table = () => {
    const cols = 1 + int(3);
    const grid: number[] = [];
    for (let c = 0; c < cols; c++) grid.push(800 + int(700));
    let rows = '';
    const count = 1 + int(8);
    for (let i = 0; i < count; i++) {
      let cells = '';
      const trPr = i === 0 && next() < 0.5 ? '<w:tblHeader/>' : next() < 0.2 ? '<w:cantSplit/>' : '';
      for (let c = 0; c < cols; c++) {
        let content = '';
        const blocks = 1 + int(3);
        for (let b = 0; b < blocks; b++) content += paragraph(true, trPr !== '');
        cells += tc(content);
      }
      rows += tr(cells, trPr);
    }
    const borders = next() < 0.5 ? '<w:tblBorders><w:top w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="12"/></w:tblBorders>' : '';
    return tbl(rows, grid, CELL_MARGINS + borders);
  };
  let body = '';
  const blocks = 20 + int(60);
  for (let i = 0; i < blocks; i++) {
    // In narrow columns a table row that cannot be divided may be taller than the page: no tables there.
    body += !columns && next() < 0.12 ? table() : paragraph(false);
    if (next() < 0.03) {
      const cols = columns && next() < 0.5 ? `<w:cols w:num="${2 + int(2)}" w:space="300"/>` : '';
      const type = pick(['', '<w:type w:val="continuous"/>', '<w:type w:val="oddPage"/>']);
      body += `<w:p><w:pPr>${sectPr(type + cols)}</w:pPr></w:p>`;
    }
  }
  body += p('end') + sectPr('<w:footerReference w:type="default" r:id="f1"/>');
  return { body, styles: STYLES, settings: COMPAT15, footers: { f1: p(field(' PAGE ') + r(' / ') + field(' NUMPAGES ')) } };
}

/** The paragraphs of the body with their paths, tables included. */
function bodyParagraphs(blocks: Block[], path: number[], out: Map<string, Paragraph>): void {
  blocks.forEach((block, i) => {
    if (block.type === 'paragraph') out.set(path.concat([i]).join('/'), block);
    else if (block.type === 'table') block.rows.forEach((row, ri) => row.cells.forEach((cell, ci) => bodyParagraphs(cell.blocks, path.concat([i, ri, ci]), out)));
  });
}

/** Every character of every paragraph of the body is on the pages exactly once, at the offset its item says. */
function assertTextComplete(doc: DocModel, pages: LayoutPage[], label: string): void {
  const paragraphs = new Map<string, Paragraph>();
  bodyParagraphs(doc.body.blocks, [], paragraphs);
  const found = new Map<string, TextItem[]>();
  for (const page of pages) {
    for (const item of texts(page)) {
      if (!item.source || item.source.story !== 'body' || item.repeat) continue;
      const key = item.source.path.join('/');
      let list = found.get(key);
      if (!list) found.set(key, (list = []));
      list.push(item);
    }
  }
  paragraphs.forEach((paragraph, key) => {
    const text = paragraphText(paragraph);
    const items = (found.get(key) || []).slice().sort((a, b) => a.source!.offset - b.source!.offset);
    let at = 0;
    for (const item of items) {
      const offset = item.source!.offset;
      // Only tabs and breaks are in the text without being in an item.
      assert.ok(offset >= at, `${label}: paragraph ${key} repeats text at ${offset}`);
      assert.ok(/^[\t\n]*$/.test(text.slice(at, offset)), `${label}: paragraph ${key} loses "${text.slice(at, offset)}"`);
      assert.equal(text.substr(offset, item.text.length), item.text, `${label}: paragraph ${key} at ${offset}`);
      at = offset + item.text.length;
    }
    assert.ok(/^[\t\n]*$/.test(text.slice(at)), `${label}: paragraph ${key} loses its end "${text.slice(at)}"`);
  });
  found.forEach((_items, key) => assert.ok(paragraphs.has(key), `${label}: text from an unknown paragraph ${key}`));
}

function assertOnPage(pages: LayoutPage[], label: string): void {
  pages.forEach((page, index) => {
    for (const item of texts(page)) {
      const where = `${label}: page ${index}, "${item.text.slice(0, 20)}" at ${item.x},${item.y}`;
      assert.ok(item.x >= -0.01 && item.x + item.width <= page.width + 0.01, where);
      assert.ok(item.y > 0 && item.y < page.height, where);
      // Body text stays between the top and bottom margins.
      if (item.source && item.source.story === 'body') assert.ok(item.y > 40 && item.y <= 360.01, where);
      assert.ok(item.width >= 0 && isFinite(item.width) && isFinite(item.x) && isFinite(item.y), where);
    }
  });
}

/** In a document of one column, the lines of top level paragraphs come one below the other in document order. */
function assertOrdered(pages: LayoutPage[], label: string): void {
  pages.forEach((page, index) => {
    let lastKey = -1;
    let lastOffset = -1;
    let lastY = -Infinity;
    for (const item of texts(page)) {
      if (!item.source || item.source.story !== 'body' || item.source.path.length !== 1) continue;
      const key = item.source.path[0];
      const sameLine = Math.abs(item.y - lastY) < 14;
      if (key < lastKey || (key === lastKey && item.source.offset < lastOffset)) assert.fail(`${label}: page ${index} is out of order at paragraph ${key}`);
      // Runs of one line may sit higher or lower (superscripts, sizes); a new line is clearly lower.
      if (!sameLine) assert.ok(item.y > lastY, `${label}: page ${index}, paragraph ${key}: a line at ${item.y} after one at ${lastY}`);
      lastKey = key;
      lastOffset = item.source.offset;
      if (!sameLine || item.y > lastY) lastY = item.y;
    }
  });
}

test('doc layout: random documents keep every character, on the page, in order', async () => {
  for (let seed = 1; seed <= 60; seed++) {
    const columns = seed % 3 === 0;
    const spec = randomSpec(seed, columns);
    const doc = await readSpec(spec);
    const layout = new DocLayout(doc, fakeMeasurer);
    assert.equal(layout.layout().done, true);
    const label = `seed ${seed}`;
    assert.ok(layout.pages.length >= 1 && layout.pages.length < 400, label);
    assertTextComplete(doc, layout.pages, label);
    assertOnPage(layout.pages, label);
    if (!columns) assertOrdered(layout.pages, label);

    // The same document in steps of three pages, and of one band at a time.
    const stepped = new DocLayout(await readSpec(spec), fakeMeasurer);
    let calls = 0;
    while (!stepped.layout({ pages: 3 }).done) assert.ok(++calls < 1000, label);
    assert.equal(JSON.stringify(stepped.pages), JSON.stringify(layout.pages), label);
    if (seed % 10 === 0) {
      const timed = new DocLayout(await readSpec(spec), fakeMeasurer);
      calls = 0;
      while (!timed.layout({ milliseconds: 0 }).done) assert.ok(++calls < 100000, label);
      assert.equal(JSON.stringify(timed.pages), JSON.stringify(layout.pages), label);
    }
  }
});

/** Lays out and checks that layout ends with a sane number of pages and finite numbers. */
async function survives(spec: DocxSpec | string, maxPages: number, measurer: TextMeasurer = fakeMeasurer): Promise<LayoutPage[]> {
  const full: DocxSpec = typeof spec === 'string' ? { body: spec } : spec;
  if (full.styles === undefined) full.styles = STYLES;
  const layout = new DocLayout(await readSpec(full), measurer);
  let calls = 0;
  while (!layout.layout({ pages: 50 }).done) assert.ok(++calls < 10000);
  assert.ok(layout.pages.length >= 1 && layout.pages.length <= maxPages, `pages: ${layout.pages.length}`);
  for (const page of layout.pages) {
    assert.ok(isFinite(page.width) && isFinite(page.height) && page.width > 0 && page.height > 0);
    for (const item of page.items) {
      if (item.type === 'line') assert.ok(isFinite(item.x1) && isFinite(item.y1) && isFinite(item.x2) && isFinite(item.y2));
      else assert.ok(isFinite(item.x) && isFinite(item.y) && isFinite(item.width), JSON.stringify(item).slice(0, 200));
    }
  }
  return layout.pages;
}

/** Characters of text with a source on the pages: what must not change however the text is broken. */
const count = (pages: LayoutPage[]) => pages.reduce((sum, page) => sum + texts(page).reduce((n, item) => n + (item.source && !item.repeat ? item.text.length : 0), 0), 0);

test('doc layout: degenerate pages and margins', async () => {
  const section = (size: string, margins: string, more = '') => `<w:sectPr><w:pgSz ${size}/><w:pgMar ${margins}/>${more}</w:sectPr>`;
  const body = filler(30);
  // No paper at all, margins larger than the paper, negative margins, a paper of one twip.
  await survives(body + section('w:w="0" w:h="0"', 'w:top="0" w:right="0" w:bottom="0" w:left="0"'), 40);
  const squeezed = await survives(body + section('w:w="6000" w:h="6000"', 'w:top="5000" w:right="5000" w:bottom="5000" w:left="5000" w:header="9000" w:footer="9000"'), 2000);
  assert.equal(count(squeezed), 80);
  await survives(body + section('w:w="6000" w:h="6000"', 'w:top="-600" w:right="600" w:bottom="-600" w:left="600"'), 40);
  await survives(body + section('w:w="1" w:h="1"', 'w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'), 2000);
  // A header and a footer taller than the page: the text keeps its room between the margins.
  const refs = '<w:headerReference w:type="default" r:id="h1"/><w:footerReference w:type="default" r:id="f1"/>';
  const tall = await survives({ body: body + sectPr(refs), headers: { h1: filler(40, 'H') }, footers: { f1: filler(40, 'F') } }, 3);
  assert.equal(tall.length, 2);
});

test('doc layout: absurd columns', async () => {
  const many = await survives(filler(30) + sectPr('<w:cols w:num="10000" w:space="720"/>'), 40);
  assert.equal(count(many), 80);
  const listed = '<w:cols w:num="3" w:equalWidth="0"><w:col w:w="0" w:space="0"/><w:col w:w="0" w:space="100000"/><w:col w:w="900000"/></w:cols>';
  assert.equal(count(await survives(p(words(40)) + sectPr(listed), 400)), words(40).length);
  await survives(filler(5) + sectPr('<w:cols w:num="0"/>'), 2);
});

test('doc layout: a chain of keep-with-next through a long document', async () => {
  let body = '';
  for (let i = 0; i < 3000; i++) body += p('K' + i, '<w:keepNext/><w:keepLines/>');
  const pages = await survives(body + sectPr(), 151);
  assert.equal(pages.length, 150);
  assert.equal(texts(pages[149]).length, 20);
  assert.equal(count(pages), 10 * 2 + 90 * 3 + 900 * 4 + 2000 * 5);
});

test('doc layout: degenerate paragraphs', async () => {
  const text = words(30);
  const body =
    p(text, '<w:ind w:left="900000" w:right="900000"/>') +
    p(text, '<w:ind w:left="-3000" w:hanging="5000"/>') +
    p(text, '<w:spacing w:line="0" w:lineRule="exact"/>') +
    p(text, '<w:spacing w:line="1000000" w:lineRule="exact"/>') +
    p(text, '<w:spacing w:before="1000000" w:after="1000000"/>') +
    p(text, '<w:spacing w:line="1000000" w:lineRule="auto"/>') +
    p(r(text, '<w:sz w:val="3000"/>')) +
    p(r(text, '<w:sz w:val="1"/>')) +
    p(r(text, '<w:spacing w:val="-2000"/>')) +
    p(r(text, '<w:w w:val="1"/>')) +
    p(r('a') + '<w:r><w:tab/><w:tab/><w:tab/></w:r>' + r('b'), '<w:tabs><w:tab w:val="right" w:pos="900000"/><w:tab w:val="center" w:pos="1"/><w:tab w:val="decimal" w:pos="31680"/></w:tabs>') +
    p('<w:r>' + new Array(3001).join('<w:tab/>') + '</w:r>') +
    p('<w:r>' + new Array(501).join('<w:br w:type="column"/>') + '</w:r>');
  const pages = await survives({ body: body + sectPr(), settings: '<w:defaultTabStop w:val="0"/>' }, 4000);
  assert.ok(pages.length > 3);
});

test('doc layout: degenerate tables', async () => {
  const row = tr(tc(words(10)) + tc(words(10)));
  // Columns of no width, a grid that does not match the cells, spans beyond the grid, a huge and a zero row height.
  await survives(tbl(row + row, [0, 0], '<w:tblW w:w="0" w:type="dxa"/><w:tblLayout w:type="fixed"/>') + sectPr(), 200);
  await survives(tbl(tr(tc('a', '<w:gridSpan w:val="10000"/>') + tc('b')) + tr(tc('c'), '<w:gridBefore w:val="5000"/><w:gridAfter w:val="5000"/>'), [1000]) + sectPr(), 5);
  await survives(tbl(tr(tc('a') + tc('b'), '<w:trHeight w:val="1000000" w:hRule="exact"/>') + tr(tc('c') + tc('d'), '<w:trHeight w:val="1000000" w:hRule="atLeast"/>'), [1000, 1000]) + p('after') + sectPr(), 6);
  await survives(tbl(tr(tc('a', '<w:tcW w:w="900000" w:type="dxa"/>')), [900000], '<w:tblInd w:w="-900000" w:type="dxa"/><w:tblLayout w:type="fixed"/>') + sectPr(), 5);
  await survives(tbl(tr(tc(p('x'), '<w:vMerge/>') + tc('b')) + tr(tc(p('y'), '<w:vMerge w:val="restart"/>') + tc('c')), [1000, 1000]) + sectPr(), 5);
  await survives(tbl('', [1000]) + tbl(tr(''), [1000]) + p('x') + sectPr(), 2);
  // A continued cell nothing started keeps its text.
  const kept = await survives(tbl(tr(tc(p('lonely'), '<w:vMerge/>')), [1500]) + sectPr(), 2);
  assert.equal(count(kept), 6);
  // Header rows as tall as the page are not repeated for ever.
  let rows = tr(tc(filler(19, 'H')), '<w:tblHeader/>');
  for (let i = 0; i < 30; i++) rows += tr(tc('R' + i));
  const headers = await survives(tbl(rows, [1500]) + sectPr(), 40);
  assert.equal(count(headers), 10 * 2 + 9 * 3 + 10 * 2 + 20 * 3);
  // Every row is a header row.
  let all = '';
  for (let i = 0; i < 50; i++) all += tr(tc('H' + i), '<w:tblHeader/>');
  assert.equal(count(await survives(tbl(all, [1500]) + sectPr(), 5)), 10 * 2 + 40 * 3);
});

test('doc layout: tables nested as deep as the reader allows', async () => {
  let inner = p('deep');
  // Find the deepest nesting the reader accepts and lay that out.
  let deepest = '';
  for (let depth = 1; depth <= 40; depth++) {
    inner = tbl(tr(tc(inner)), [3000]);
    try {
      await readSpec({ body: inner + sectPr() });
      deepest = inner;
    } catch {
      break;
    }
  }
  assert.ok(deepest !== '');
  const pages = await survives(deepest + sectPr(), 400);
  assert.equal(count(pages), 4);
});

test('doc layout: a measurer that returns nonsense', async () => {
  const body = p(words(30)) + tbl(tr(tc(words(5)) + tc('b')), [1500, 1500]) + sectPr();
  const cases: TextMeasurer[] = [
    { width: () => NaN, metrics: () => ({ ascent: NaN, descent: NaN, lineGap: NaN }) },
    { width: () => Infinity, metrics: () => ({ ascent: Infinity, descent: 1, lineGap: -Infinity }) },
    { width: () => -5, metrics: () => ({ ascent: -1, descent: -1, lineGap: -1 }) },
    { width: () => 0, metrics: () => ({ ascent: 0, descent: 0, lineGap: 0 }) },
    { width: () => 1e9, metrics: () => ({ ascent: 1e9, descent: 1e9, lineGap: 1e9 }) },
  ];
  for (const measurer of cases) await survives(body, 500, measurer);
});

test('doc layout: drawings of absurd sizes and positions', async () => {
  const drawing = (extent: string, position: string) =>
    `<w:r><w:drawing><wp:anchor distT="2147483647" distB="2147483647" distL="2147483647" distR="2147483647" simplePos="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:posOffset>${position}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${position}</wp:posOffset></wp:positionV>` +
    `<wp:extent ${extent}/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="d"/><a:graphic><a:graphicData uri="x"/></a:graphic></wp:anchor></w:drawing></w:r>`;
  const body =
    p(drawing('cx="2147483647" cy="2147483647"', '0') + r(words(40))) +
    p(drawing('cx="0" cy="0"', '-2147483647') + r(words(40))) +
    p(drawing('cx="3000000" cy="100000"', '0') + r(words(40))) +
    filler(30);
  const pages = await survives(body + sectPr(), 300);
  assert.equal(count(pages), 3 * words(40).length + 80);
});

test('doc layout: layout stops for good at the page limit', async () => {
  let body = '';
  for (let i = 0; i < 100020; i++) body += '<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>';
  const layout = new DocLayout(await readSpec({ body }), fakeMeasurer);
  let calls = 0;
  while (!layout.layout({ pages: 20000 }).done) assert.ok(++calls < 10);
  assert.equal(layout.pages.length, 100000);
  assert.equal(layout.truncated, true);
  assert.deepEqual(layout.layout(), { done: true, updated: [] });
});

test('doc layout: speed on a long document', async (t) => {
  const parts: string[] = [];
  for (let i = 0; i < 50000; i++) parts.push(p(words(8 + (i % 60), 3 + (i % 5), i)));
  const doc = await readSpec({ body: parts.join('') + sectPr().replace('w:w="6000" w:h="6000"', 'w:w="11906" w:h="16838"'), styles: STYLES });
  const layout = new DocLayout(doc, fakeMeasurer);
  let start = Date.now();
  layout.layout({ pages: 10 });
  const first = Date.now() - start;
  assert.equal(layout.pages.length, 10);
  start = Date.now();
  while (layout.pages.length < 1000 && !layout.layout({ pages: 100 }).done);
  const thousand = first + (Date.now() - start);
  start = Date.now();
  layout.layout();
  const total = thousand + (Date.now() - start);
  t.diagnostic(`50,000 paragraphs: first 10 pages ${first} ms, 1,000 pages ${thousand} ms, all ${layout.pages.length} pages ${total} ms`);
  // Measured: about 10 ms, 250 ms and 0.8 s. The limits leave room for a slow machine.
  assert.ok(first < 500, `first pages took ${first} ms`);
  assert.ok(thousand < 5000, `1,000 pages took ${thousand} ms`);
  assert.ok(total < 30000, `the whole document took ${total} ms`);
  let lines = 0;
  for (const page of layout.pages) lines += page.items.length;
  assert.ok(lines > 50000);
});
