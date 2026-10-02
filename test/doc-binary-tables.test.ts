// Tables of a Word binary file: paragraphs gathered into rows and cells, the
// grid made from the cell edges, merged cells, nested tables, and what the
// row marks say about borders, shading and heights.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { DocLayout } from '../src/layout/doc/layout';
import type { Block, DocModel, Table } from '../src/model/doc/types';
import { fakeMeasurer, texts } from './doc-layout-fixtures';
import { DocBuilder, FC, IN_TABLE, ROW_END, brc80, cat, cellTexts, defTable, paragraphs, sprm, stsh, tablesOf, tc80, textOf, u16le, u32le, u8s } from './doc-binary-fixtures';

/**
 * A document from rows of cell texts. Each cell is a paragraph that ends with
 * a cell mark; each row ends with a row mark that carries `rowSprms(index)`.
 * "\r" inside a cell text starts another paragraph of the cell.
 */
function tableDoc(rows: string[][], rowSprms: (row: number) => Uint8Array[], before = 'before\r', after = 'after\r'): DocBuilder {
  const builder = new DocBuilder();
  let text = before;
  const marks: { from: number; to: number; row?: number }[] = [];
  rows.forEach((cells, row) => {
    for (const cell of cells) {
      // Every paragraph of the cell, the last one ending with the cell mark.
      const parts = cell.split('\r');
      parts.forEach((part, i) => {
        const from = text.length;
        text += part + (i === parts.length - 1 ? '\u0007' : '\r');
        marks.push({ from, to: text.length });
      });
    }
    const from = text.length;
    text += '\u0007';
    marks.push({ from, to: text.length, row });
  });
  const tableEnd = text.length;
  text += after;
  builder.text(text);
  if (before) builder.papx(0, before.length, 0);
  for (const mark of marks) {
    if (mark.row === undefined) builder.papx(mark.from, mark.to, 0, IN_TABLE);
    else builder.papx(mark.from, mark.to, 0, IN_TABLE, ROW_END, ...rowSprms(mark.row));
  }
  if (after) builder.papx(tableEnd, text.length, 0);
  return builder;
}

const only = (model: DocModel): Table => {
  const tables = tablesOf(model.body.blocks);
  assert.equal(tables.length, 1);
  return tables[0];
};

test('paragraphs in a table become rows and cells; the grid comes from the cell edges', () => {
  const model = tableDoc(
    [
      ['A1', 'B1'],
      ['A2', 'B2'],
    ],
    () => [sprm(0x9602, u16le(108)), defTable([-108, 2000, 5000])],
  ).read();
  assert.deepEqual(
    model.body.blocks.map((block) => block.type),
    ['paragraph', 'table', 'paragraph'],
  );
  const table = only(model);
  assert.deepEqual(cellTexts(table), [
    ['A1', 'B1'],
    ['A2', 'B2'],
  ]);
  assert.deepEqual(table.grid, [2108, 3000]);
  assert.deepEqual(table.rows[0].cells.map((cell) => cell.props.width), [
    { type: 'dxa', value: 2108 },
    { type: 'dxa', value: 3000 },
  ]);
  // The edge of the table is 108 twips left of the margin: its text starts at the margin.
  assert.deepEqual(table.props.indent, { type: 'dxa', value: 0 });
  assert.deepEqual(table.props.cellMargins, { left: { type: 'dxa', value: 108 }, right: { type: 'dxa', value: 108 } });
  assert.equal(table.rows[0].cells[0].props.gridSpan, undefined);
  assert.deepEqual(textOf(model.body.blocks), ['before', 'after']);
});

test('a cell may hold several paragraphs', () => {
  const table = only(tableDoc([['one\rtwo\rthree', 'x']], () => [defTable([0, 1000, 2000])]).read());
  assert.deepEqual(cellTexts(table), [['one|two|three', 'x']]);
});

test('rows with different edges share a grid: cells span columns, rows skip columns', () => {
  const edges = [
    [0, 1000, 2000, 3000],
    [0, 2000, 3000],
    [1000, 3000],
    [0, 1000, 2500],
  ];
  const table = only(tableDoc([['a', 'b', 'c'], ['ab', 'c'], ['bc'], ['a', 'b+']], (row) => [defTable(edges[row])]).read());
  assert.deepEqual(table.grid, [1000, 1000, 500, 500]);
  const spans = table.rows.map((row) => row.cells.map((cell) => cell.props.gridSpan || 1));
  assert.deepEqual(spans, [[1, 1, 2], [2, 2], [3], [1, 2]]);
  assert.equal(table.rows[2].props.gridBefore, 1);
  assert.deepEqual(table.rows[2].props.widthBefore, { type: 'dxa', value: 1000 });
  assert.equal(table.rows[3].props.gridAfter, 1);
  assert.equal(table.rows[0].props.gridBefore, undefined);
});

test('horizontally merged cells become one cell that spans their columns', () => {
  // The first cell starts the merge, the second and third are merged into it.
  const cells = [tc80(0x0001), tc80(0x0002), tc80(0x0002), tc80()];
  const table = only(tableDoc([['merged', '', '', 'd']], () => [defTable([0, 1000, 2000, 3000, 4000], cells)]).read());
  assert.deepEqual(cellTexts(table), [['merged', 'd']]);
  assert.equal(table.rows[0].cells[0].props.gridSpan, 3);
  assert.deepEqual(table.rows[0].cells[0].props.width, { type: 'dxa', value: 3000 });
  assert.equal(table.rows[0].cells[0].props.hMerge, undefined);
});

test('text a merged-away cell still holds is kept with the cell it is merged into', () => {
  const cells = [tc80(0x0001), tc80(0x0002)];
  const table = only(tableDoc([['kept', 'also kept']], () => [defTable([0, 1000, 2000], cells)]).read());
  assert.deepEqual(cellTexts(table), [['kept|also kept']]);
});

test('vertically merged cells: the first restarts, the ones below continue', () => {
  const def = (flags: number) => defTable([0, 1000, 2000], [tc80(flags), tc80()]);
  const table = only(tableDoc([['top', 'x'], ['', 'y'], ['', 'z'], ['own', 'w']], (row) => [def(row === 0 ? 0x0060 : row < 3 ? 0x0020 : 0)]).read());
  assert.deepEqual(
    table.rows.map((row) => row.cells[0].props.vMerge),
    ['restart', 'continue', 'continue', undefined],
  );
  // sprmTVertMerge says the same for one cell at a time.
  const bySprm = only(tableDoc([['top', 'x'], ['', 'y']], (row) => [defTable([0, 1000, 2000]), sprm(0xd62b, 2, 1, row === 0 ? 3 : 1)]).read());
  assert.deepEqual(
    bySprm.rows.map((row) => row.cells[1].props.vMerge),
    ['restart', 'continue'],
  );
});

test('borders: of the cells, of the table, and set for a range of cells', () => {
  const single = brc80(4, 1, 1);
  const thick = brc80(12, 1, 6);
  const cells = [tc80(0, [single, single, 0, 0xffffffff]), tc80()];
  const table = only(
    tableDoc([['a', 'b']], () => [
      defTable([0, 1000, 2000], cells),
      // The borders of the table, in Word 97's form: top, left, bottom, right, inside horizontal, inside vertical.
      sprm(0xd605, 24, u32le(thick), u32le(thick), u32le(thick), u32le(thick), u32le(single), u32le(single)),
      // A double bottom border for the second cell.
      sprm(0xd620, 7, 1, 2, 0x04, u32le(brc80(6, 3))),
    ]).read(),
  );
  const [a, b] = table.rows[0].cells;
  const line = { style: 'single', size: 4, space: 0, color: { rgb: '000000' } };
  // A side the cell says nothing about is left to the borders of the table; "no border" is said as such.
  assert.deepEqual(a.props.borders, { top: line, left: line, right: { style: 'nil' } });
  assert.deepEqual(b.props.borders, { bottom: { style: 'double', size: 6, space: 0, color: { auto: true } } });
  assert.equal(table.props.borders!.top!.size, 12);
  assert.deepEqual(table.props.borders!.top!.color, { rgb: 'FF0000' });
  assert.equal(table.props.borders!.insideV!.size, 4);
});

test('the borders of Word 2000 and later carry a full colour', () => {
  const brc = (r: number, g: number, b: number, width: number, type: number) => u8s(r, g, b, 0, width, type, 0, 0);
  const table = only(
    tableDoc([['a']], () => [
      defTable([0, 1000]),
      sprm(0xd613, 48, brc(1, 2, 3, 8, 1), brc(0, 0, 0, 8, 1), brc(0, 0, 0, 8, 1), brc(0, 0, 0, 8, 1), brc(0, 0, 0, 4, 1), brc(0, 0, 0, 4, 1)),
      sprm(0xd62f, 11, 0, 1, 0x01 | 0x08, brc(0xaa, 0xbb, 0xcc, 16, 3)),
    ]).read(),
  );
  assert.deepEqual(table.props.borders!.top, { style: 'single', size: 8, space: 0, color: { rgb: '010203' } });
  const border = { style: 'double', size: 16, space: 0, color: { rgb: 'AABBCC' } };
  assert.deepEqual(table.rows[0].cells[0].props.borders, { top: border, right: border });
});

test('shading, vertical alignment, row height, header rows and rows that stay whole', () => {
  const table = only(
    tableDoc([['a', 'b'], ['c', 'd']], (row) =>
      row === 0
        ? [
            defTable([0, 1000, 2000], [tc80(0x0080), tc80(0x0100)]),
            sprm(0xd609, 4, u16le((1 << 10) | (7 << 5) | 1), u16le(0xffff)),
            sprm(0x9407, u16le(400)),
            sprm(0x3404, 1),
            sprm(0x3403, 1),
            sprm(0x5400, u16le(1)),
          ]
        : [defTable([0, 1000, 2000]), sprm(0x9407, u16le(0x10000 - 300)), sprm(0xd62c, 3, 0, 2, 2), sprm(0xd62d, 12, 0, 1, u8s(0xff, 0, 0, 0, 0, 0xff, 0, 0), u16le(0))],
    ).read(),
  );
  const [first, second] = table.rows;
  assert.deepEqual(first.cells.map((cell) => cell.props.vAlign), ['center', 'bottom']);
  assert.deepEqual(first.cells[0].props.shading, { pattern: 'solid', color: { rgb: '000000' }, fill: { rgb: 'FFFF00' } });
  assert.equal(first.cells[1].props.shading, undefined);
  assert.deepEqual(first.props, { height: 400, heightRule: 'atLeast', header: true, cantSplit: true, align: 'center' });
  assert.equal(table.props.align, 'center');
  assert.deepEqual(second.props, { height: 300, heightRule: 'exact' });
  assert.deepEqual(second.cells.map((cell) => cell.props.vAlign), ['bottom', 'bottom']);
  assert.deepEqual(second.cells[0].props.shading, { pattern: 'clear', color: { rgb: 'FF0000' }, fill: { rgb: '00FF00' } });
});

test('cell margins and widths set by sprm; what a later row says differently goes with that row', () => {
  const margins = (value: number) => sprm(0xd634, 6, 0, 1, 0x02 | 0x08, 3, u16le(value));
  const table = only(
    tableDoc([['a', 'b'], ['c', 'd']], (row) => [
      defTable([0, 1000, 2000]),
      margins(row === 0 ? 70 : 140),
      sprm(0xd632, 6, 1, 2, 0x01, 3, u16le(55)),
      sprm(0xf614, 2, u16le(5000)),
      sprm(0x3615, 0),
    ]).read(),
  );
  assert.deepEqual(table.props.cellMargins, { left: { type: 'dxa', value: 70 }, right: { type: 'dxa', value: 70 } });
  assert.deepEqual(table.props.indent, { type: 'dxa', value: 70 });
  assert.deepEqual(table.props.width, { type: 'pct', value: 5000 });
  assert.equal(table.props.layout, 'fixed');
  assert.deepEqual(table.rows[0].cells[1].props.margins, { top: { type: 'dxa', value: 55 } });
  assert.equal(table.rows[0].tableProps, undefined);
  assert.deepEqual(table.rows[1].tableProps, { cellMargins: { left: { type: 'dxa', value: 140 }, right: { type: 'dxa', value: 140 } } });
});

test('the indent of a table counts the cell margin in effect: the one of its style when the rows state none', () => {
  const read = (rowSprms: Uint8Array[]) => {
    const builder = tableDoc([['a']], () => [defTable([-70, 1000]), sprm(0x563a, u16le(1)), ...rowSprms]);
    builder.table(FC.Stshf, stsh([{ sti: 0, name: 'Normal' }, { kind: 3, name: 'Tight', tapx: sprm(0xd634, 6, 0, 1, 2 | 8, 3, u16le(70)) }]));
    return only(builder.read());
  };
  const styled = read([]);
  assert.equal(styled.props.style, 'Tight');
  assert.equal(styled.props.cellMargins, undefined);
  // The edge is 70 twips left of the margin and the margin of the style is 70: the text starts at the margin.
  assert.deepEqual(styled.props.indent, { type: 'dxa', value: 0 });
  // A margin of the row's own wins over the style.
  assert.deepEqual(read([sprm(0x9602, u16le(30))]).props.indent, { type: 'dxa', value: -40 });
  // Without a style and without margins: what Word assumes.
  assert.deepEqual(only(tableDoc([['a']], () => [defTable([-70, 1000])]).read()).props.indent, { type: 'dxa', value: 38 });
});

test('a row may be made of cell inserts instead of a table definition, and its sprms may lie in the Data stream', () => {
  const builder = new DocBuilder();
  const text = 'a\u0007b\u0007c\u0007\u0007';
  builder.text(text);
  builder.papx(0, 2, 0, IN_TABLE).papx(2, 4, 0, IN_TABLE).papx(4, 6, 0, IN_TABLE);
  // The left edge, then three cells of 1000, 2500 and 500 twips, one insert each; the third is widened afterwards.
  const row = cat(
    IN_TABLE,
    ROW_END,
    sprm(0x9601, u16le(0x10000 - 100)),
    sprm(0x7621, 0, 1, u16le(1000)),
    sprm(0x7621, 1, 1, u16le(2500)),
    sprm(0x7621, 2, 1, u16le(500)),
    sprm(0x7623, 2, 3, u16le(800)),
  );
  const at = builder.dataBlob(cat(u16le(row.length), row));
  // The paragraph itself only says where its properties are (sprmPTableProps).
  builder.papx(6, 7, 0, sprm(0x646b, u32le(at)));
  const table = only(builder.read());
  assert.deepEqual(cellTexts(table), [['a', 'b', 'c']]);
  assert.deepEqual(table.grid, [1000, 2500, 800]);
});

test('cells deleted from a row definition take their width with them', () => {
  const table = only(tableDoc([['a', 'c']], () => [defTable([0, 1000, 2000, 3500]), sprm(0x5622, 1, 2)]).read());
  assert.deepEqual(table.grid, [1000, 1500]);
});

test('a table in a table: a deeper level of paragraphs inside a cell', () => {
  const builder = new DocBuilder();
  // outer cell 1: "o1" + inner table (2 cells) + "tail"; outer cell 2: "o2".
  const text = 'o1\ri1\ri2\r\rtail\u0007o2\u0007\u0007end\r';
  builder.text(text);
  const depth = (n: number) => sprm(0x6649, u32le(n));
  builder.papx(0, 3, 0, IN_TABLE, depth(1));
  builder.papx(3, 6, 0, IN_TABLE, depth(2), sprm(0x244b, 1));
  builder.papx(6, 9, 0, IN_TABLE, depth(2), sprm(0x244b, 1));
  builder.papx(9, 10, 0, IN_TABLE, depth(2), sprm(0x244b, 1), sprm(0x244c, 1), defTable([0, 700, 1400]));
  builder.papx(10, 15, 0, IN_TABLE, depth(1));
  builder.papx(15, 18, 0, IN_TABLE, depth(1));
  builder.papx(18, 19, 0, IN_TABLE, depth(1), ROW_END, defTable([0, 3000, 6000]));
  builder.papx(19, 23, 0);
  const model = builder.read();
  const outer = only(model);
  assert.equal(outer.rows.length, 1);
  const [first, second] = outer.rows[0].cells;
  assert.deepEqual(first.blocks.map((block) => block.type), ['paragraph', 'table', 'paragraph']);
  assert.deepEqual(textOf(first.blocks), ['o1', 'tail']);
  const inner = first.blocks[1] as Table;
  assert.deepEqual(cellTexts(inner), [['i1', 'i2']]);
  assert.deepEqual(inner.grid, [700, 700]);
  assert.deepEqual(textOf(second.blocks), ['o2']);
  assert.deepEqual(textOf(model.body.blocks), ['end']);
});

test('tables that follow each other without text between them are one table; text between them parts them', () => {
  const builder = new DocBuilder();
  const text = 'a\u0007\u0007between\rb\u0007\u0007';
  builder.text(text);
  builder.papx(0, 2, 0, IN_TABLE).papx(2, 3, 0, IN_TABLE, ROW_END, defTable([0, 1000]));
  builder.papx(3, 11, 0);
  builder.papx(11, 13, 0, IN_TABLE).papx(13, 14, 0, IN_TABLE, ROW_END, defTable([0, 1000]));
  const model = builder.read();
  assert.deepEqual(model.body.blocks.map((block) => block.type), ['table', 'paragraph', 'table']);
});

test('cells without a row mark after them are still a row; a row mark without edges gives a table without a grid', () => {
  const builder = new DocBuilder();
  const text = 'a\u0007b\u0007';
  builder.text(text);
  builder.papx(0, 2, 0, IN_TABLE).papx(2, 4, 0, IN_TABLE);
  const cut = only(builder.read());
  assert.deepEqual(cellTexts(cut), [['a', 'b']]);
  assert.deepEqual(cut.grid, []);

  const bare = only(tableDoc([['a', 'b']], () => []).read());
  assert.deepEqual(cellTexts(bare), [['a', 'b']]);
  assert.deepEqual(bare.grid, []);
  assert.equal(bare.props.indent, undefined);
});

test('a row with more cells than its definition, or fewer, keeps every cell', () => {
  const more = only(tableDoc([['a', 'b', 'c']], () => [defTable([0, 1000, 2000])]).read());
  assert.deepEqual(cellTexts(more), [['a', 'b', 'c']]);
  assert.deepEqual(more.rows[0].cells.map((cell) => cell.props.width && cell.props.width.value), [1000, 1000, undefined]);
  const fewer = only(tableDoc([['a']], () => [defTable([0, 1000, 2000, 3000])]).read());
  assert.deepEqual(cellTexts(fewer), [['a']]);
  assert.deepEqual(fewer.grid, [1000]);
});

test('a paragraph that claims an absurd table depth does not make the reader build one', () => {
  const builder = new DocBuilder();
  const text = 'deep\rnext\r';
  builder.text(text);
  builder.papx(0, 5, 0, IN_TABLE, sprm(0x6649, u32le(0x7fffffff)));
  builder.papx(5, 10, 0);
  const model = builder.read();
  // Nested as deep as the reader goes, and no deeper; the text is all there.
  let depth = 0;
  let blocks: Block[] = model.body.blocks;
  for (;;) {
    const table = tablesOf(blocks)[0];
    if (!table) break;
    depth++;
    blocks = table.rows[0].cells[0].blocks;
  }
  assert.equal(depth, 32);
  assert.deepEqual(textOf(blocks), ['deep']);
  assert.deepEqual(textOf(model.body.blocks), ['next']);
});

test('edges that go backwards are pulled into order', () => {
  const table = only(tableDoc([['a', 'b', 'c']], () => [defTable([0, 2000, 1000, 3000])]).read());
  // The cell whose edges ran backwards is a twip wide: every cell has a grid column of its own.
  assert.deepEqual(table.grid, [2000, 1, 999]);
  assert.deepEqual(table.rows[0].cells.map((cell) => cell.props.gridSpan || 1), [1, 1, 1]);
});

test('a floating table says where it sits', () => {
  const table = only(
    tableDoc([['a']], () => [defTable([0, 1000]), sprm(0x360d, (1 << 4) | (2 << 6)), sprm(0x940e, u16le(0x10000 - 4)), sprm(0x940f, u16le(500)), sprm(0x9410, u16le(180)), sprm(0x941e, u16le(190))]).read(),
  );
  assert.deepEqual(table.props.floating, { vertAnchor: 'page', horzAnchor: 'page', xAlign: 'center', y: 500, leftFromText: 180, rightFromText: 190 });
});

test('the layout engine lays the table out: cells side by side in the widths of the grid', () => {
  const model = tableDoc(
    [
      ['A1', 'B1'],
      ['A2', 'B2'],
    ],
    () => [sprm(0x9602, u16le(150)), defTable([-150, 1350, 4350])],
  ).read();
  const layout = new DocLayout(model, fakeMeasurer);
  assert.equal(layout.layout().done, true);
  assert.equal(layout.pages.length, 1);
  const items = texts(layout.pages[0]);
  const find = (text: string) => items.find((item) => item.text === text)!;
  assert.deepEqual(items.map((item) => item.text), ['before', 'A1', 'B1', 'A2', 'B2', 'after']);
  // The text of the first column starts at the margin, the second column 1500 twips (100 px) further right.
  assert.equal(find('A1').x, find('before').x);
  assert.equal(find('B1').x - find('A1').x, 100);
  assert.equal(find('A2').x, find('A1').x);
  assert.ok(find('A2').y > find('A1').y);
  assert.equal(find('B1').y, find('A1').y);
  assert.deepEqual(paragraphs(model.body.blocks).length, 2);
});
