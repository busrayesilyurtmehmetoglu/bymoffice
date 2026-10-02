// Tables, hyperlinks and pictures.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { DocLayout } from '../src/layout/doc/layout';
import type { DrawingItem, Run, Table } from '../src/model/doc/types';
import { serializeXml } from '../src/xml';
import { DocHistory, deleteColumn, deleteRow, insertColumn, insertHyperlink, insertImage, insertRow, insertTable, insertText, type DocRange } from '../src/ops/doc/index';
import { checkInvariants, checkLayout, measurer, open, p, paragraphAt, pos, r, range, run, shapeAt, tbl, tc, texts, tr } from './doc-ops-fixtures';

const history = (doc: Parameters<typeof texts>[0]) => new DocHistory<DocRange | undefined>(doc);
const table = (doc: Parameters<typeof texts>[0], index: number) => doc.body.blocks[index] as Table;
/** The cells of a table as rows of "text" or "text*span", vertical merges marked "^" (restart) and "|" (continue). */
const grid = (t: Table) =>
  t.rows.map((row) =>
    row.cells
      .map((cell) => {
        const text = cell.blocks.map((block) => (block.type === 'paragraph' ? block.children.map((inline) => (inline.type === 'run' ? inline.content.map((item) => (item.type === 'text' ? item.text : '')).join('') : '')).join('') : '[t]')).join('/');
        return text + (cell.props.gridSpan ? '*' + cell.props.gridSpan : '') + (cell.props.vMerge === 'restart' ? '^' : cell.props.vMerge === 'continue' ? '|' : '');
      })
      .join(' '),
  );

test('insertTable makes an even grid over the text width, with borders of its own or the grid style', async () => {
  const doc = await open(p('onetwo') + p('end'));
  const h = history(doc);
  const after = run(h, insertTable(doc, pos(0, 3), 2, 3));
  assert.deepEqual(doc.body.blocks.map((block) => block.type), ['paragraph', 'table', 'paragraph', 'paragraph']);
  assert.deepEqual(texts(doc), ['one', '', '', '', '', '', '', 'two', 'end']);
  assert.deepEqual(after.focus, pos([1, 0, 0, 0], 0));
  const t = table(doc, 1);
  // The page is 9000 twips wide with margins of 600.
  assert.deepEqual(t.grid, [2600, 2600, 2600]);
  assert.deepEqual(t.rows[1].cells[2].props.width, { type: 'dxa', value: 2600 });
  assert.equal(t.props.style, undefined);
  assert.equal(t.props.borders!.insideV!.style, 'single');
  assert.equal(t.srcStart, undefined);
  checkInvariants(doc);
  assert.ok(checkLayout(doc) > 0);

  const styled = await open({ body: p('x'), styles: '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>' });
  run(history(styled), insertTable(styled, pos(0, 1), 1, 1));
  assert.equal(table(styled, 1).props.style, 'TableGrid');
  assert.equal(table(styled, 1).props.borders, undefined);
  // A table does not end the story: a paragraph follows it.
  assert.deepEqual(styled.body.blocks.map((block) => block.type), ['paragraph', 'table', 'paragraph']);
  // Nor does it touch another table.
  run(history(styled), insertTable(styled, pos(2, 0), 1, 2));
  run(history(styled), insertTable(styled, pos(0, 1), 1, 2));
  assert.deepEqual(styled.body.blocks.map((block) => block.type), ['paragraph', 'table', 'paragraph', 'table', 'paragraph', 'table', 'paragraph']);
  // In a cell the table takes the width of the cell.
  run(history(styled), insertTable(styled, pos([1, 0, 0, 0], 0), 2, 2));
  const inner = table(styled, 1).rows[0].cells[0].blocks[0] as Table;
  assert.equal(inner.type, 'table');
  assert.equal(inner.grid.length, 2);
  assert.ok(inner.grid[0] * 2 < table(styled, 1).grid[0]);
  checkInvariants(styled);
  for (const [rows, columns] of [[0, 1], [1, 0], [1, 64], [1.5, 2], [100000, 2]]) assert.throws(() => insertTable(styled, pos(0, 0), rows, columns), RangeError);
});

const MERGED = tbl(
  tr(tc('a1') + tc('b1', '<w:gridSpan w:val="2"/>') + tc('d1', '<w:vMerge w:val="restart"/>')) + tr(tc('a2') + tc('b2') + tc('c2') + tc('', '<w:vMerge/>')) + tr(tc('a3') + tc('b3') + tc('c3') + tc('d3')),
  [1000, 1000, 1000, 1000],
);

test('rows: inserted with the cells of the row beside them, continuing vertical merges; deleted with the merge moving on', async () => {
  const doc = await open(p('x') + MERGED + p('y'));
  const h = history(doc);
  assert.deepEqual(grid(table(doc, 1)), ['a1 b1*2 d1^', 'a2 b2 c2 |', 'a3 b3 c3 d3']);
  let after = run(h, insertRow(doc, pos([1, 0, 1, 0], 1), 'below'));
  assert.deepEqual(grid(table(doc, 1)), ['a1 b1*2 d1^', ' *2 |', 'a2 b2 c2 |', 'a3 b3 c3 d3']);
  assert.deepEqual(after.focus, pos([1, 1, 1, 0], 0));
  after = run(h, insertRow(doc, pos([1, 0, 0, 0], 0), 'above'));
  assert.deepEqual(grid(table(doc, 1)), [' *2 ', 'a1 b1*2 d1^', ' *2 |', 'a2 b2 c2 |', 'a3 b3 c3 d3']);
  assert.deepEqual(after.focus, pos([1, 0, 0, 0], 0));
  run(h, insertRow(doc, pos([1, 4, 3, 0], 0), 'below'));
  assert.deepEqual(grid(table(doc, 1))[5], '   ');
  checkInvariants(doc);
  // The row that begins the merged cell goes: the cell below it begins the merge now.
  after = run(h, deleteRow(doc, pos([1, 1, 2, 0], 0)));
  assert.deepEqual(grid(table(doc, 1)), [' *2 ', ' *2 ^', 'a2 b2 c2 |', 'a3 b3 c3 d3', '   ']);
  assert.deepEqual(after.focus, pos([1, 1, 2, 0], 0));
  checkInvariants(doc);
  assert.ok(checkLayout(doc) > 0);
  while (h.undo());
  assert.deepEqual(grid(table(doc, 1)), ['a1 b1*2 d1^', 'a2 b2 c2 |', 'a3 b3 c3 d3']);
  assert.throws(() => insertRow(doc, pos(0, 0), 'below'), RangeError);
  assert.throws(() => deleteRow(doc, pos(2, 0)), RangeError);
});

test('the last row takes the table with it, and a paragraph stays where one is needed', async () => {
  const doc = await open(tbl(tr(tc('only')), [1000]) + p('after'));
  const h = history(doc);
  let after = run(h, deleteRow(doc, pos([0, 0, 0, 0], 2)));
  assert.deepEqual(doc.body.blocks.map((block) => block.type), ['paragraph']);
  assert.deepEqual(after.focus, pos(0, 0));
  h.undo();
  // A table that is all a cell holds leaves an empty paragraph.
  const nested = await open(tbl(tr(tc(tbl(tr(tc('inner')), [500]) + p(''))), [1000]) + p('after'));
  after = run(history(nested), deleteColumn(nested, pos([0, 0, 0, 0, 0, 0, 0], 0)));
  assert.deepEqual(texts(nested), ['', 'after']);
  assert.deepEqual(after.focus, pos([0, 0, 0, 0], 0));
  const alone = await open(p('before') + tbl(tr(tc('only')), [1000]));
  run(history(alone), deleteRow(alone, pos([1, 0, 0, 0], 0)));
  assert.deepEqual(texts(alone), ['before']);
});

test('columns: a cell in every row, or a wider merged cell; deleting narrows merged cells', async () => {
  const doc = await open(p('x') + MERGED + p('y'));
  const h = history(doc);
  let after = run(h, insertColumn(doc, pos([1, 1, 1, 0], 0), 'right'));
  // The new column is between b and c, inside the cell that spans both in the first row.
  assert.deepEqual(grid(table(doc, 1)), ['a1 b1*3 d1^', 'a2 b2  c2 |', 'a3 b3  c3 d3']);
  assert.deepEqual(table(doc, 1).grid, [1000, 1000, 1000, 1000, 1000]);
  assert.deepEqual(after.focus, pos([1, 1, 2, 0], 0));
  after = run(h, insertColumn(doc, pos([1, 0, 0, 0], 0), 'left'));
  assert.deepEqual(grid(table(doc, 1)), [' a1 b1*3 d1^', ' a2 b2  c2 |', ' a3 b3  c3 d3']);
  assert.deepEqual(after.focus, pos([1, 0, 0, 0], 0));
  // Right of the vertically merged column: the new cells are merged the same way.
  run(h, insertColumn(doc, pos([1, 0, 3, 0], 0), 'right'));
  assert.deepEqual(grid(table(doc, 1)), [' a1 b1*3 d1^ ^', ' a2 b2  c2 | |', ' a3 b3  c3 d3 ']);
  checkInvariants(doc);
  assert.ok(checkLayout(doc) > 0);
  // Deleting the column of b2: the merged cell above it loses one column.
  after = run(h, deleteColumn(doc, pos([1, 1, 2, 0], 0)));
  assert.deepEqual(grid(table(doc, 1)), [' a1 b1*2 d1^ ^', ' a2  c2 | |', ' a3  c3 d3 ']);
  assert.equal(table(doc, 1).grid.length, 6);
  // Deleting the columns of the merged cell takes both of them.
  run(h, deleteColumn(doc, pos([1, 0, 2, 0], 0)));
  assert.deepEqual(grid(table(doc, 1)), [' a1 d1^ ^', ' a2 | |', ' a3 d3 ']);
  assert.equal(table(doc, 1).grid.length, 4);
  checkInvariants(doc);
  assert.ok(checkLayout(doc) > 0);
  while (h.undo());
  assert.deepEqual(grid(table(doc, 1)), ['a1 b1*2 d1^', 'a2 b2 c2 |', 'a3 b3 c3 d3']);
  // The last column takes the table.
  const single = await open(tbl(tr(tc('a')) + tr(tc('b')), [1000]) + p('after'));
  run(history(single), deleteColumn(single, pos([0, 1, 0, 0], 0)));
  assert.deepEqual(texts(single), ['after']);
});

test('table width follows its columns when it is stated in twips', async () => {
  const doc = await open('<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="1200"/><w:gridCol w:w="800"/></w:tblGrid>' + tr(tc('a', '<w:tcW w:w="1200" w:type="dxa"/>') + tc('b', '<w:tcW w:w="800" w:type="dxa"/>')) + '</w:tbl>' + p(''));
  const h = history(doc);
  run(h, insertColumn(doc, pos([0, 0, 1, 0], 0), 'right'));
  assert.deepEqual(table(doc, 0).grid, [1200, 800, 800]);
  assert.equal(table(doc, 0).props.width!.value, 2800);
  assert.deepEqual(table(doc, 0).rows[0].cells[2].props.width, { type: 'dxa', value: 800 });
  run(h, deleteColumn(doc, pos([0, 0, 0, 0], 0)));
  assert.deepEqual(table(doc, 0).grid, [800, 800]);
  assert.equal(table(doc, 0).props.width!.value, 1600);
});

test('a hyperlink gets a relationship of the story, and undoing takes it away', async () => {
  const doc = await open({ body: p(r('see ') + r('the site', '<w:b/>') + r(' now')), rels: [['rId7', 'hyperlink', 'https://old.example/', true]] });
  const h = history(doc);
  const relationships = doc.body.relationships;
  const size = relationships.size;
  const after = run(h, insertHyperlink(doc, range(pos(0, 4), pos(0, 12)), 'https://example.org/a?b=1&c=2'));
  assert.deepEqual(after, range(pos(0, 4), pos(0, 12)));
  assert.equal(doc.body.relationships.size, size + 1);
  const link = (paragraphAt(doc, 0).children[1] as Run).wrap!;
  assert.equal(link.kind, 'hyperlink');
  const rel = doc.body.relationships.get((link as { rId: string }).rId)!;
  assert.deepEqual(rel, { id: 'rId8', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', target: 'https://example.org/a?b=1&c=2', external: true, part: null });
  assert.equal((link as { target: string }).target, rel.target);
  // Without a "Hyperlink" character style the text is coloured and underlined directly.
  assert.deepEqual(shapeAt(doc, 0), [':see ', 'bu:the site@hyperlink', ': now']);
  assert.deepEqual((paragraphAt(doc, 0).children[1] as Run).props.color, { rgb: '0563C1' });
  // The layout shows the link.
  const layout = new DocLayout(doc, measurer);
  layout.layout();
  assert.ok(layout.pages[0].items.some((item) => item.type === 'text' && item.link === 'https://example.org/a?b=1&c=2'));
  h.undo();
  assert.equal(doc.body.relationships, relationships, 'the very table of relationships is back');
  assert.equal(doc.body.relationships.size, size);
});

test('hyperlinks: to a bookmark, over an existing link, with a style, with a caret', async () => {
  const doc = await open({
    body: p(r('a') + `<w:hyperlink r:id="rId7">${r('old')}</w:hyperlink>` + '<w:ins w:id="1" w:author="A"><w:r><w:t>new</w:t></w:r></w:ins>' + r('z')) + p('second'),
    styles: '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:u w:val="single"/></w:rPr></w:style>',
    rels: [['rId7', 'hyperlink', 'https://old.example/', true]],
  });
  const h = history(doc);
  const size = doc.body.relationships.size;
  run(h, insertHyperlink(doc, range(pos(0, 0), pos(0, 8)), '#mark'));
  assert.equal(doc.body.relationships.size, size, 'a link into the document needs no relationship');
  // One link around all of it: the old link is gone and the tracked insertion is inside the new one.
  assert.deepEqual(shapeAt(doc, 0), ['{Hyperlink}:aold@hyperlink', '{Hyperlink}:new@ins@hyperlink', '{Hyperlink}:z@hyperlink']);
  const links = new Set(paragraphAt(doc, 0).children.map((inline) => { let w = inline.wrap; while (w && w.kind !== 'hyperlink') w = w.parent; return w; }));
  assert.equal(links.size, 1);
  assert.equal((Array.from(links)[0] as { anchor: string }).anchor, 'mark');
  // A caret gets the address as its text.
  const after = run(h, insertHyperlink(doc, range(pos(1, 3), pos(1, 3)), 'mailto:someone@example.org'));
  assert.deepEqual(shapeAt(doc, 1), [':sec', '{Hyperlink}:mailto:someone@example.org@hyperlink', ':ond']);
  assert.deepEqual(after, range(pos(1, 3), pos(1, 29)));
  assert.throws(() => insertHyperlink(doc, range(pos(0, 0), pos(1, 1)), 'https://example.org/'), RangeError);
  assert.throws(() => insertHyperlink(doc, range(pos(0, 0), pos(0, 1)), ''), RangeError);
  assert.throws(() => insertHyperlink(doc, range(pos(0, 0), pos(0, 1)), 'a\nb'), RangeError);
  assert.equal(doc.body.relationships.size, size + 1);
});

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]);

test('a picture: bytes in doc.media under a new part name, a relationship, and a drawing the layout shows', async () => {
  const doc = await open({ body: p('ab') + p('cd'), headers: { rIdH: p('header') }, parts: [['word/media/image4.png', PNG]], rels: [['rId3', 'image', 'media/image4.png']] });
  const h = history(doc);
  assert.equal(doc.media, undefined);
  const relationships = doc.body.relationships;
  const after = run(h, insertImage(doc, pos(0, 1), { bytes: PNG, type: 'image/png', width: 952500, height: 476250, description: 'a <picture>' }));
  assert.deepEqual(after.focus, pos(0, 1));
  assert.deepEqual(Array.from(doc.media!.keys()), ['/word/media/image5.png']);
  assert.equal(doc.media!.get('/word/media/image5.png')!.bytes, PNG);
  assert.deepEqual(shapeAt(doc, 0), [':a<drawing>b']);
  const drawing = (paragraphAt(doc, 0).children[0] as Run).content[1] as DrawingItem;
  const rel = doc.body.relationships.get(drawing.image!.rId)!;
  assert.deepEqual(rel, { id: rel.id, type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', target: 'media/image5.png', external: false, part: '/word/media/image5.png' });
  assert.equal(drawing.image!.target, '/word/media/image5.png');
  assert.equal(drawing.width, 952500);
  // The kept XML is a complete inline picture that names the relationship.
  const xml = serializeXml({ declaration: null, leading: [], root: drawing.xml, trailing: [] });
  assert.ok(xml.indexOf(`r:embed="${rel.id}"`) > 0);
  assert.ok(xml.indexOf('<wp:extent cx="952500" cy="476250"/>') > 0);
  assert.equal(/descr="([^"]*)"/.exec(xml)![1].replace('&gt;', '>'), 'a &lt;picture>');
  assert.ok(/^<w:drawing><wp:inline xmlns:wp=/.test(xml));
  // The layout places it: 100 x 50 px.
  const layout = new DocLayout(doc, measurer);
  layout.layout();
  const image = layout.pages[0].items.find((item) => item.type === 'image');
  assert.ok(image && image.type === 'image' && image.width === 100 && image.height === 50 && image.part === '/word/media/image5.png');
  // A second picture, in the header: another part name, a relationship of the header part, another drawing id.
  run(h, insertImage(doc, pos(0, 0, 'header:rIdH'), { bytes: PNG, type: 'image/jpeg', width: 9525, height: 9525 }));
  assert.deepEqual(Array.from(doc.media!.keys()), ['/word/media/image5.png', '/word/media/image6.jpeg']);
  const header = doc.headers.get('rIdH')!;
  assert.equal(header.relationships.size, 1);
  const second = (paragraphAt(doc, 0, 'header:rIdH').children[0] as Run).content[0] as DrawingItem;
  assert.notEqual(second.xml.children.length, 0);
  const ids = [drawing, second].map((item) => serializeXml({ declaration: null, leading: [], root: item.xml, trailing: [] }).match(/docPr id="(\d+)"/)![1]);
  assert.notEqual(ids[0], ids[1]);
  // Typing after the picture stays after it.
  run(h, insertText(doc, pos(0, 1), 'X'));
  assert.deepEqual(shapeAt(doc, 0), [':a<drawing>Xb']);
  while (h.undo());
  assert.equal(doc.media, undefined, 'the map that was not there is gone again');
  assert.equal(doc.body.relationships, relationships);
  assert.equal(header.relationships.size, 0);
  for (const bad of [{ bytes: PNG, type: 'image/svg+xml', width: 1, height: 1 }, { bytes: new Uint8Array(0), type: 'image/png', width: 1, height: 1 }, { bytes: PNG, type: 'image/png', width: 0, height: 1 }, { bytes: [1] as unknown as Uint8Array, type: 'image/png', width: 1, height: 1 }]) {
    assert.throws(() => insertImage(doc, pos(0, 0), bad), RangeError);
  }
  assert.equal(doc.media, undefined);
});
