// Tables of an RTF file: paragraphs gathered into rows and cells, the grid
// made from the cell edges, merged cells, nested tables, and what the row
// descriptions say about borders, padding, shading and heights.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { resolveCellProps, resolveTableProps } from '../src/formats/docx/resolve';
import { DocLayout } from '../src/layout/doc/layout';
import { paragraphText, type Block, type DocModel, type Table } from '../src/model/doc/types';
import { fakeMeasurer, lines, texts as pageTexts } from './doc-layout-fixtures';
import { doc, paragraphsOf, read, tables, texts } from './rtf-fixtures';

/** The texts of the cells of a table, row by row; a nested table shows as "[...]". */
function cellTexts(table: Table): string[][] {
  return table.rows.map((row) =>
    row.cells.map((cell) =>
      cell.blocks
        .map((block: Block) => (block.type === 'paragraph' ? paragraphText(block) : block.type === 'table' ? '[' + cellTexts(block).map((cells) => cells.join(',')).join(';') + ']' : '?'))
        .join('|'),
    ),
  );
}

const first = (model: DocModel): Table => tables(model)[0];

test('rows and cells: the grid, the widths and where the table starts', () => {
  const body = String.raw`before\par\trowd\trgaph108\trleft-108\cellx2000\cellx5000\pard\intbl A\cell B\cell\row\trowd\trgaph108\trleft-108\cellx2000\cellx5000\pard\intbl C\cell D\cell\row\pard after\par`;
  const model = read(doc(body));
  assert.deepEqual(model.body.blocks.map((block) => block.type), ['paragraph', 'table', 'paragraph']);
  const table = first(model);
  assert.deepEqual(cellTexts(table), [
    ['A', 'B'],
    ['C', 'D'],
  ]);
  assert.deepEqual(table.grid, [2108, 3000]);
  assert.deepEqual(table.rows[0].cells.map((cell) => cell.props.width), [
    { type: 'dxa', value: 2108 },
    { type: 'dxa', value: 3000 },
  ]);
  // The indent is measured to the text of the first cell: the edge plus the padding.
  assert.deepEqual(table.props.indent, { type: 'dxa', value: 0 });
  assert.deepEqual(table.props.cellMargins, { left: { type: 'dxa', value: 108 }, right: { type: 'dxa', value: 108 } });
  assert.equal(table.rows[1].tableProps, undefined);
});

test('the description in effect at "\\row" counts, wherever it stands', () => {
  // Word writes the description before the cells and again after them; here the two differ.
  const word = String.raw`\trowd\cellx1000\cellx2000\pard\intbl A\cell B\cell\pard\intbl{\trowd\trrh400\cellx3000\cellx7000\row}\pard x\par`;
  assert.deepEqual(first(read(doc(word))).grid, [3000, 4000]);
  assert.equal(first(read(doc(word))).rows[0].props.height, 400);
  // Other writers state it once, before the cells, and keep it for the rows that follow.
  const once = String.raw`\trowd\cellx3000\cellx7000\pard\intbl A\cell B\cell\row\pard\intbl C\cell D\cell\row\pard x\par`;
  const table = first(read(doc(once)));
  assert.deepEqual(table.grid, [3000, 4000]);
  assert.deepEqual(cellTexts(table), [
    ['A', 'B'],
    ['C', 'D'],
  ]);
});

test('a cell holds every paragraph up to its mark; an empty cell holds an empty paragraph', () => {
  const body = String.raw`\trowd\cellx3000\cellx6000\cellx9000\pard\intbl one\par two\cell\cell three\cell\row\pard x\par`;
  const table = first(read(doc(body)));
  assert.deepEqual(cellTexts(table), [['one|two', '', 'three']]);
  assert.equal(table.rows[0].cells[1].blocks.length, 1);
});

test('row properties: height, header row, keeping together, alignment', () => {
  const body =
    String.raw`\trowd\trrh-300\trhdr\trkeep\trqc\cellx3000\pard\intbl A\cell\row` +
    String.raw`\trowd\trrh500\trqc\cellx3000\pard\intbl B\cell\row\trowd\trqc\cellx3000\pard\intbl C\cell\row\pard x\par`;
  const table = first(read(doc(body)));
  assert.deepEqual(table.rows[0].props, { height: 300, heightRule: 'exact', header: true, cantSplit: true, align: 'center' });
  assert.deepEqual(table.rows[1].props, { height: 500, heightRule: 'atLeast', align: 'center' });
  assert.deepEqual(table.rows[2].props, { align: 'center' });
  assert.equal(table.props.align, 'center');
});

test('padding: "\\trpadd" with its unit, else half the gap, else none', () => {
  const padded = first(read(doc(String.raw`\trowd\trgaph70\trleft-100\trpaddl100\trpaddr120\trpaddt30\trpaddfl3\trpaddfr3\trpaddft3\cellx3000\pard\intbl A\cell\row\pard x\par`)));
  assert.deepEqual(padded.props.cellMargins, { left: { type: 'dxa', value: 100 }, top: { type: 'dxa', value: 30 }, right: { type: 'dxa', value: 120 } });
  assert.deepEqual(padded.props.indent, { type: 'dxa', value: 0 });
  // A unit without a value is a padding of nothing.
  const zero = first(read(doc(String.raw`\trowd\trleft-50\trpaddfl3\trpaddfr3\cellx3000\pard\intbl A\cell\row\pard x\par`)));
  assert.deepEqual(zero.props.cellMargins, { left: { type: 'dxa', value: 0 }, right: { type: 'dxa', value: 0 } });
  assert.deepEqual(zero.props.indent, { type: 'dxa', value: -50 });
  const none = first(read(doc(String.raw`\trowd\cellx3000\pard\intbl A\cell\row\pard x\par`)));
  assert.deepEqual(none.props.cellMargins, { left: { type: 'dxa', value: 0 }, right: { type: 'dxa', value: 0 } });
  assert.deepEqual(none.props.indent, { type: 'dxa', value: 0 });
});

test('cells merged across and down', () => {
  const body =
    String.raw`\trowd\clmgf\cellx2000\clmrg\cellx4000\clvmgf\cellx6000\pard\intbl wide\cell\cell top\cell\row` +
    String.raw`\trowd\cellx2000\cellx4000\clvmrg\cellx6000\pard\intbl a\cell b\cell\cell\row\pard x\par`;
  const table = first(read(doc(body)));
  assert.deepEqual(table.grid, [2000, 2000, 2000]);
  assert.deepEqual(cellTexts(table), [
    ['wide', 'top'],
    ['a', 'b', ''],
  ]);
  const [wide, top] = table.rows[0].cells;
  assert.equal(wide.props.gridSpan, 2);
  assert.deepEqual(wide.props.width, { type: 'dxa', value: 4000 });
  assert.equal(wide.props.hMerge, undefined);
  assert.equal(top.props.vMerge, 'restart');
  assert.equal(table.rows[1].cells[2].props.vMerge, 'continue');
});

test('rows with different edges share one grid', () => {
  const body =
    String.raw`\trowd\cellx3000\cellx6000\pard\intbl A\cell B\cell\row` +
    String.raw`\trowd\cellx2000\cellx6000\pard\intbl C\cell D\cell\row` +
    String.raw`\trowd\trleft1000\cellx5000\pard\intbl E\cell\row\pard x\par`;
  const table = first(read(doc(body)));
  assert.deepEqual(table.grid, [1000, 1000, 1000, 2000, 1000]);
  assert.deepEqual(table.rows[0].cells.map((cell) => cell.props.gridSpan), [3, 2]);
  assert.deepEqual(table.rows[1].cells.map((cell) => cell.props.gridSpan), [2, 3]);
  const last = table.rows[2];
  assert.equal(last.props.gridBefore, 1);
  assert.deepEqual(last.props.widthBefore, { type: 'dxa', value: 1000 });
  assert.equal(last.props.gridAfter, 1);
  assert.equal(last.cells[0].props.gridSpan, 3);
});

test('borders of the table and of cells, with widths and colours', () => {
  const body =
    String.raw`\trowd\trbrdrt\brdrs\brdrw10 \trbrdrl\brdrs\brdrw10 \trbrdrb\brdrs\brdrw10 \trbrdrr\brdrs\brdrw10 \trbrdrh\brdrdot\brdrw5 \trbrdrv\brdrdb\brdrw15\brdrcf1 ` +
    String.raw`\clbrdrt\brdrs\brdrw30\brdrcf2 \clbrdrl\brdrnone \clbrdrb\brdrtbl \clbrdrr\brdrdash\brdrw20 \cellx3000\clbrdrt\brdrnil\cellx6000\pard\intbl A\cell B\cell\row\pard x\par`;
  const model = read(doc(body));
  const table = first(model);
  const thin = { style: 'single', size: 4 };
  assert.deepEqual(table.props.borders, { top: thin, left: thin, bottom: thin, right: thin, insideH: { style: 'dotted', size: 2 }, insideV: { style: 'double', size: 6, color: { rgb: 'FF0000' } } });
  assert.deepEqual(table.rows[0].cells[0].props.borders, {
    top: { style: 'single', size: 12, color: { rgb: '0000FF' } },
    left: { style: 'none' },
    right: { style: 'dashed', size: 8 },
  });
  assert.deepEqual(table.rows[0].cells[1].props.borders, { top: { style: 'nil' } });
  assert.deepEqual(resolveTableProps(model, table).borders!.insideH, { style: 'dotted', size: 2 });
});

test('a border word after a paragraph border belongs to the paragraph, not to the row', () => {
  const body = String.raw`\trowd\clbrdrt\brdrs\brdrw10 \cellx3000\pard\intbl\brdrb\brdrdb\brdrw20 A\cell\row\pard x\par`;
  const table = first(read(doc(body)));
  assert.deepEqual(table.rows[0].cells[0].props.borders, { top: { style: 'single', size: 4 } });
  assert.deepEqual(paragraphsOf(table.rows[0].cells[0].blocks)[0].props.borders, { bottom: { style: 'double', size: 8 } });
});

test('shading, vertical alignment, direction and padding of cells', () => {
  const body =
    String.raw`\trowd\clvertalc\clcbpat3\cltxbtlr\cellx2000\clvertalb\clshdng2500\clcfpat1\clcbpat2\cellx4000\clpadt150\clpadft3\clpadl40\clpadfl3\clpadr60\clpadfr3\clNoWrap\clFitText\cellx6000` +
    String.raw`\pard\intbl A\cell B\cell C\cell\row\pard x\par`;
  const model = read(doc(body));
  const [a, b, c] = first(model).rows[0].cells;
  assert.equal(a.props.vAlign, 'center');
  assert.deepEqual(a.props.shading, { pattern: 'clear', fill: { rgb: 'FFFF00' } });
  assert.equal(a.props.textDirection, 'btLr');
  assert.equal(b.props.vAlign, 'bottom');
  assert.deepEqual(b.props.shading, { pattern: 'pct25', fill: { rgb: '0000FF' }, color: { rgb: 'FF0000' } });
  // Word writes the left padding as "\clpadt" and the top padding as "\clpadl".
  assert.deepEqual(c.props.margins, { left: { type: 'dxa', value: 150 }, top: { type: 'dxa', value: 40 }, right: { type: 'dxa', value: 60 } });
  assert.equal(c.props.noWrap, true);
  assert.equal(c.props.fitText, true);
  assert.equal(resolveCellProps(model, { table: first(model), row: 0, cell: 0 }).vAlign, 'center');
});

test('a row that states other borders or padding than the first carries them itself', () => {
  const body =
    String.raw`\trowd\trgaph100\trbrdrt\brdrs\brdrw10 \cellx3000\pard\intbl A\cell\row` +
    String.raw`\trowd\trgaph100\trbrdrt\brdrs\brdrw10 \cellx3000\pard\intbl B\cell\row` +
    String.raw`\trowd\trgaph50\cellx3000\pard\intbl C\cell\row\pard x\par`;
  const table = first(read(doc(body)));
  assert.equal(table.rows[1].tableProps, undefined);
  assert.deepEqual(table.rows[2].tableProps, { borders: { top: { style: 'nil' } }, cellMargins: { left: { type: 'dxa', value: 50 }, right: { type: 'dxa', value: 50 } } });
});

test('width, layout and direction of the table', () => {
  const body = String.raw`\trowd\rtlrow\trautofit1\trftsWidth3\trwWidth5000\trspdl20\trspdfl3\cellx5000\pard\intbl A\cell\row\pard x\par`;
  const table = first(read(doc(body)));
  assert.deepEqual(table.props.width, { type: 'dxa', value: 5000 });
  assert.equal(table.props.layout, 'autofit');
  assert.equal(table.props.bidiVisual, true);
  assert.deepEqual(table.props.cellSpacing, { type: 'dxa', value: 20 });
  const pct = first(read(doc(String.raw`\trowd\trftsWidth2\trwWidth2500\cellx5000\pard\intbl A\cell\row\pard x\par`)));
  assert.deepEqual(pct.props.width, { type: 'pct', value: 2500 });
});

test('a table that text flows around', () => {
  const body = String.raw`\trowd\tphmrg\tpvpara\tposx500\tposnegy-200\tdfrmtxtLeft141\tdfrmtxtRight142\tabsnoovrlp1\cellx3000\pard\intbl A\cell\row\pard x\par`;
  const table = first(read(doc(body)));
  assert.deepEqual(table.props.floating, { horzAnchor: 'margin', vertAnchor: 'text', x: 500, y: -200, leftFromText: 141, rightFromText: 142 });
  assert.equal(table.props.overlap, 'never');
  const aligned = first(read(doc(String.raw`\trowd\tphpg\tpvpg\tposxc\tposyb\cellx3000\pard\intbl A\cell\row\pard x\par`)));
  assert.deepEqual(aligned.props.floating, { horzAnchor: 'page', vertAnchor: 'page', xAlign: 'center', yAlign: 'bottom' });
});

test('a table in a cell, as Word writes it', () => {
  const body =
    String.raw`\trowd\cellx4000\cellx8000\pard\intbl outer A\par` +
    String.raw`\pard\intbl\itap2 in 1\nestcell{\nonesttables\par}\pard\intbl\itap2 in 2\nestcell{\nonesttables\par}` +
    String.raw`\pard\intbl\itap2{\*\nesttableprops\trowd\trgaph30\cellx1500\cellx3000\nestrow}{\nonesttables\par}` +
    String.raw`\pard\intbl\itap2 in 3\nestcell in 4\nestcell\pard\intbl\itap2{\*\nesttableprops\trowd\trgaph30\cellx1500\cellx3000\nestrow}{\nonesttables\par}` +
    String.raw`\pard\intbl tail\cell\pard\intbl outer B\cell\pard\intbl{\trowd\cellx4000\cellx8000\row}\pard after\par`;
  const model = read(doc(body));
  const table = first(model);
  assert.deepEqual(cellTexts(table), [['outer A|[in 1,in 2;in 3,in 4]|tail', 'outer B']]);
  assert.deepEqual(table.grid, [4000, 4000]);
  const inner = table.rows[0].cells[0].blocks[1] as Table;
  assert.deepEqual(inner.grid, [1500, 1500]);
  assert.deepEqual(inner.props.cellMargins, { left: { type: 'dxa', value: 30 }, right: { type: 'dxa', value: 30 } });
  assert.deepEqual(texts(model), ['outer A', 'in 1', 'in 2', 'in 3', 'in 4', 'tail', 'outer B', 'after']);
});

test('the description of a nested row does not replace that of the row around it', () => {
  // The outer row is described once, before its cells; the nested one in between must not change it.
  const body =
    String.raw`\trowd\cellx4000\cellx8000\pard\intbl\itap2 in\nestcell\pard\intbl\itap2{\*\nesttableprops\trowd\cellx1000\nestrow}` +
    String.raw`\pard\intbl A\cell B\cell\row\pard after\par`;
  const table = first(read(doc(body)));
  assert.deepEqual(table.grid, [4000, 4000]);
  assert.deepEqual(cellTexts(table), [['[in]|A', 'B']]);
});

test('two tables with a paragraph between them stay two; rows without one are one table', () => {
  const row = (text: string) => String.raw`\trowd\cellx3000\pard\intbl ` + text + String.raw`\cell\row`;
  const model = read(doc(row('A') + row('B') + String.raw`\pard between\par` + row('C') + String.raw`\pard after\par`));
  assert.deepEqual(model.body.blocks.map((block) => block.type), ['table', 'paragraph', 'table', 'paragraph']);
  assert.equal(tables(model)[0].rows.length, 2);
  assert.equal(tables(model)[1].rows.length, 1);
});

test('"\\itap" alone puts a paragraph into a table; "\\pard" takes the next one out', () => {
  const model = read(doc(String.raw`\trowd\cellx3000\pard\itap1 A\cell\row\pard out\par`));
  assert.deepEqual(model.body.blocks.map((block) => block.type), ['table', 'paragraph']);
});

test('cells without the end of their row, and a row mark without cells', () => {
  const cut = read(doc(String.raw`\trowd\cellx3000\cellx6000\pard\intbl A\cell B\cell`));
  assert.deepEqual(cellTexts(first(cut)), [['A', 'B']]);
  const open = read(doc(String.raw`\trowd\cellx3000\pard\intbl A\par B`));
  assert.deepEqual(cellTexts(first(open)), [['A|B']]);
  const stray = read(doc(String.raw`a\par\row\row b\par`));
  assert.deepEqual(stray.body.blocks.map((block) => block.type), ['paragraph', 'paragraph']);
  // Text before a row mark without a cell mark is a cell.
  const unmarked = read(doc(String.raw`\trowd\cellx3000\pard\intbl A\row\pard x\par`));
  assert.deepEqual(cellTexts(first(unmarked)), [['A']]);
});

test('more cells than the row describes, and edges that go backwards', () => {
  const extra = first(read(doc(String.raw`\trowd\cellx3000\pard\intbl A\cell B\cell C\cell\row\pard x\par`)));
  assert.deepEqual(cellTexts(extra), [['A', 'B', 'C']]);
  assert.deepEqual(extra.grid, [3000]);
  const backwards = first(read(doc(String.raw`\trowd\cellx3000\cellx2000\cellx5000\pard\intbl A\cell B\cell C\cell\row\pard x\par`)));
  assert.deepEqual(backwards.grid, [3000, 1, 1999]);
});

test('a table lays out inside the page, its first edge where "\\trleft" says', () => {
  // A page of 12240 twips with margins of 1800: the text area starts at 120 px. The row starts 108 twips left of it.
  const body = String.raw`\trowd\trgaph108\trleft-108\trbrdrl\brdrs\brdrw10 \trbrdrr\brdrs\brdrw10 \cellx2892\cellx5892\pard\intbl\fs24 Left\cell Right\cell\row\pard x\par`;
  const model = read(doc(body));
  const layout = new DocLayout(model, fakeMeasurer);
  assert.equal(layout.layout().done, true);
  const page = layout.pages[0];
  const left = pageTexts(page).find((item) => item.text === 'Left')!;
  const right = pageTexts(page).find((item) => item.text === 'Right')!;
  // The text of the first cell starts at the margin; that of the second 3000 twips further right.
  assert.equal(Math.round(left.x), 120);
  assert.equal(Math.round(right.x), 320);
  const xs = lines(page).filter((line) => line.x1 === line.x2).map((line) => Math.round(line.x1 * 10) / 10);
  assert.deepEqual(xs, [112.8, 512.8]);
});
