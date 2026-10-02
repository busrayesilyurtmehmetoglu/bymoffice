// What a review of the doc reader found: tables of writers that build rows
// from inserted cells, where a table starts, pictures scaled beyond what a
// page holds, toggles of character styles, and the smaller things beside them.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { resolveRunProps } from '../src/formats/docx/resolve';
import { numberFormat } from '../src/formats/doc/sprm';
import type { DocModel, DrawingItem, Paragraph, Run, Table } from '../src/model/doc/types';
import { DocBuilder, FC, IN_TABLE, ROW_END, cat, cellTexts, defTable, paragraphs, plc, runsOf, sprm, stsh, tablesOf, textOf, u16le, u32le, u8s } from './doc-binary-fixtures';

// ---------------------------------------------------------------------------
// Tables

/** A document that is one table: each cell a paragraph, each row ended by a row mark with `rowSprms(row)`. */
function tableDoc(rows: string[][], rowSprms: (row: number) => Uint8Array[]): DocBuilder {
  const builder = new DocBuilder();
  let text = '';
  const marks: { from: number; to: number; row?: number }[] = [];
  rows.forEach((cells, row) => {
    for (const cell of cells) {
      const from = text.length;
      text += cell + '\u0007';
      marks.push({ from, to: text.length });
    }
    const from = text.length;
    text += '\u0007';
    marks.push({ from, to: text.length, row });
  });
  const end = text.length;
  text += 'after\r';
  builder.text(text);
  for (const mark of marks) {
    if (mark.row === undefined) builder.papx(mark.from, mark.to, 0, IN_TABLE);
    else builder.papx(mark.from, mark.to, 0, IN_TABLE, ROW_END, ...rowSprms(mark.row));
  }
  builder.papx(end, text.length, 0);
  return builder;
}

const only = (model: DocModel): Table => {
  const tables = tablesOf(model.body.blocks);
  assert.equal(tables.length, 1);
  return tables[0];
};

/** sprmTInsert: `count` cells of `width` twips at the index `first`. */
const insert = (first: number, count: number, width: number) => sprm(0x7621, first, count, u16le(width));
/** sprmTCellWidth for the cells [first, limit): a width with its unit (3 twips, 2 fiftieths of a percent). */
const cellWidth = (first: number, limit: number, width: number, unit = 3) => sprm(0xd635, 5, first, limit, unit, u16le(width));
const dxaLeft = (value: number) => sprm(0x9601, u16le(value & 0xffff));
const gapHalf = (value: number) => sprm(0x9602, u16le(value));
const widthIndent = (value: number, unit = 3) => sprm(0xf661, unit, u16le(value & 0xffff));

test('rows made of inserted cells: the width a cell states itself is its width, not the width of the insert', () => {
  // Every insert carries the width of the whole table; the cells say how wide they are.
  const table = only(tableDoc([['a', 'b']], () => [insert(0, 1, 3869), insert(1, 1, 6044), cellWidth(0, 1, 1330), cellWidth(1, 2, 2077)]).read());
  assert.deepEqual(table.grid, [1330, 2077]);
  assert.deepEqual(table.rows[0].cells.map((cell) => cell.props.width), [
    { type: 'dxa', value: 1330 },
    { type: 'dxa', value: 2077 },
  ]);
  assert.deepEqual(cellTexts(table), [['a', 'b']]);

  // The inserts add up to 10760, the cells to the 10700 the table says it is wide.
  const top = only(tableDoc([['a', 'b', 'c']], () => [sprm(0xf614, 3, u16le(10700)), insert(0, 3, 3587), insert(3, 0, 0), cellWidth(0, 2, 3500), cellWidth(2, 3, 3700)]).read());
  assert.deepEqual(top.grid, [3500, 3500, 3700]);
  assert.deepEqual(top.props.width, { type: 'dxa', value: 10700 });

  // A cell that states no width, or one in another unit, keeps the width of its insert; what follows moves with it.
  const mixed = only(tableDoc([['a', 'b', 'c']], () => [insert(0, 1, 1000), insert(1, 1, 2000), insert(2, 1, 3000), cellWidth(0, 1, 400), cellWidth(2, 3, 2500, 2)]).read());
  assert.deepEqual(mixed.grid, [400, 2000, 3000]);
  assert.deepEqual(mixed.rows[0].cells[2].props.width, { type: 'dxa', value: 3000 });

  // A width of nothing is not a width.
  assert.deepEqual(only(tableDoc([['a']], () => [insert(0, 1, 1500), cellWidth(0, 1, 0)]).read()).grid, [1500]);

  // The edges of a table definition are where Word laid the cells out: a preferred width does not move them.
  const defined = only(tableDoc([['a', 'b']], () => [defTable([0, 1000, 3000]), cellWidth(0, 2, 700)]).read());
  assert.deepEqual(defined.grid, [1000, 2000]);
});

test('a nested table of inserted cells is as wide as its cells say, not as the table around it', () => {
  const builder = new DocBuilder();
  // Outer cell 1 holds an inner table of one cell; outer cell 2 is plain.
  const text = 'in\r\r\u0007o2\u0007\u0007end\r';
  builder.text(text);
  const depth = (n: number) => sprm(0x6649, u32le(n));
  builder.papx(0, 3, 0, IN_TABLE, depth(2), sprm(0x244b, 1));
  // The insert has the width of the outer table, the cell its own.
  builder.papx(3, 4, 0, IN_TABLE, depth(2), sprm(0x244b, 1), sprm(0x244c, 1), insert(0, 1, 9828), cellWidth(0, 1, 1260));
  builder.papx(4, 5, 0, IN_TABLE, depth(1));
  builder.papx(5, 8, 0, IN_TABLE, depth(1));
  builder.papx(8, 9, 0, IN_TABLE, depth(1), ROW_END, insert(0, 1, 9828), insert(1, 1, 9828), cellWidth(0, 1, 1260), cellWidth(1, 2, 8568));
  builder.papx(9, 13, 0);
  const outer = only(builder.read());
  assert.deepEqual(outer.grid, [1260, 8568]);
  const inner = outer.rows[0].cells[0].blocks[0] as Table;
  assert.equal(inner.type, 'table');
  assert.deepEqual(cellTexts(inner), [['in']]);
  assert.deepEqual(inner.grid, [1260]);
  // The inner table fits the cell it is in.
  const width = (table: Table) => table.grid.reduce((sum, column) => sum + column, 0);
  assert.ok(width(inner) <= (outer.rows[0].cells[0].props.width as { value: number }).value);
});

test('the same table written as a table definition and as inserted cells: the same indent and the same grid', () => {
  // A table indented by 93 twips with cells of 4551 and 4253 twips and cell margins of 108.
  const word = only(tableDoc([['a', 'b']], () => [dxaLeft(93), gapHalf(108), defTable([-15, 4536, 8789]), widthIndent(93)]).read());
  const inserted = only(tableDoc([['a', 'b']], () => [dxaLeft(93), gapHalf(108), insert(0, 1, 4551), insert(1, 1, 4253)]).read());
  for (const table of [word, inserted]) {
    assert.deepEqual(table.props.indent, { type: 'dxa', value: 93 });
    assert.deepEqual(table.grid, [4551, 4253]);
    assert.deepEqual(table.props.cellMargins, { left: { type: 'dxa', value: 108 }, right: { type: 'dxa', value: 108 } });
  }
  // The order of the sprms of the row does not matter: the gap may come after the cells, the place last of all.
  const late = only(tableDoc([['a', 'b']], () => [insert(0, 1, 4551), insert(1, 1, 4253), gapHalf(108), dxaLeft(93)]).read());
  assert.deepEqual(late.props.indent, { type: 'dxa', value: 93 });
  assert.deepEqual(late.grid, [4551, 4253]);
  // Word 97 states no indent: the place of the first text is the indent.
  const old = only(tableDoc([['a', 'b']], () => [dxaLeft(93), gapHalf(108), defTable([-15, 4536, 8789])]).read());
  assert.deepEqual(old.props.indent, { type: 'dxa', value: 93 });
  // The indent a row states wins over what follows from its edges.
  const stated = only(tableDoc([['a']], () => [dxaLeft(93), gapHalf(108), defTable([-15, 4536]), widthIndent(120)]).read());
  assert.deepEqual(stated.props.indent, { type: 'dxa', value: 120 });
  // An indent in another unit is not one.
  const percent = only(tableDoc([['a']], () => [dxaLeft(93), gapHalf(108), defTable([-15, 4536]), widthIndent(120, 2)]).read());
  assert.deepEqual(percent.props.indent, { type: 'dxa', value: 93 });
});

test('sprmTDxaLeft places a row of inserted cells: its first edge is half a gap to the left', () => {
  // Two rows that start at different places share a grid.
  const table = only(tableDoc([['a', 'b'], ['c', 'd']], (row) => (row === 0 ? [dxaLeft(500), gapHalf(108), insert(0, 2, 1000)] : [dxaLeft(800), gapHalf(108), insert(0, 1, 700), insert(1, 1, 1000)])).read());
  assert.deepEqual(table.grid, [300, 700, 1000]);
  assert.deepEqual(table.props.indent, { type: 'dxa', value: 500 });
  assert.equal(table.rows[0].props.gridBefore, undefined);
  assert.equal(table.rows[1].props.gridBefore, 1);
  assert.deepEqual(table.rows[1].props.widthBefore, { type: 'dxa', value: 300 });

  // The first row further in than the second: the indent is that of the leftmost row.
  const stepped = only(tableDoc([['a', 'b'], ['c', 'd']], (row) => [dxaLeft(row === 0 ? 800 : 500), gapHalf(108), insert(0, 2, 1000)]).read());
  assert.deepEqual(stepped.props.indent, { type: 'dxa', value: 500 });
  assert.equal(stepped.rows[0].props.gridBefore, 1);

  // Without a gap the edge is where the text starts; a negative place is left of the margin.
  const plain = only(tableDoc([['a']], () => [dxaLeft(200), insert(0, 1, 1000)]).read());
  assert.deepEqual(plain.props.indent, { type: 'dxa', value: 200 });
  const left = only(tableDoc([['a']], () => [dxaLeft(-300), gapHalf(70), insert(0, 1, 1000)]).read());
  assert.deepEqual(left.props.indent, { type: 'dxa', value: -300 });
  assert.deepEqual(left.grid, [1000]);

  // A table definition after it says where the edges are: they stay.
  const defined = only(tableDoc([['a'], ['b']], (row) => [dxaLeft(40), gapHalf(100), defTable(row === 0 ? [-60, 1000] : [0, 1000])]).read());
  assert.deepEqual(defined.grid, [60, 1000]);
  assert.deepEqual(defined.props.indent, { type: 'dxa', value: 40 });
  // One before it is moved by it: the text of the first cell starts where it says.
  const moved = only(tableDoc([['a'], ['b']], (row) => (row === 0 ? [defTable([0, 1000]), gapHalf(100), dxaLeft(40)] : [defTable([0, 1000])])).read());
  assert.deepEqual(moved.grid, [60, 940, 60]);
});

// ---------------------------------------------------------------------------
// Text boxes

/** An OfficeArt record: version, instance, type and content. */
const art = (version: number, instance: number, type: number, ...content: Uint8Array[]): Uint8Array => {
  const body = cat(...content);
  return cat(u16le(version | (instance << 4)), u16le(type), u32le(body.length), body);
};

/** A shape with one simple property: the 1-based number of the text of its box. */
const boxShape = (id: number, textNumber: number): Uint8Array =>
  art(0xf, 0, 0xf004, art(2, 75, 0xf00a, u32le(id), u32le(0xa00)), art(3, 1, 0xf00b, cat(u16le(0x0080), u32le(textNumber << 16))));

const spa = (id: number): Uint8Array => cat(u32le(id), u32le(0), u32le(0), u32le(3000), u32le(1000), u16le(3 << 5), u32le(0));

const drawingsOf = (paragraph: Paragraph): DrawingItem[] =>
  runsOf(paragraph)
    .flatMap((run: Run) => run.content)
    .filter((item): item is DrawingItem => item.type === 'drawing');

/**
 * A document with text boxes. `body` and the texts of the boxes may hold drawing characters
 * (U+0008); `shapes[i]` is the number of the box text the i-th of them shows, counted through
 * the body and then the boxes.
 */
function boxDoc(body: string, boxes: string[], shapes: number[]): DocBuilder {
  const builder = new DocBuilder();
  // The text of each box ends with a paragraph mark; one more mark closes the list of them.
  const boxText = boxes.join('') + '\r';
  const text = body + boxText + '\r';
  builder.text(text, { compressed: true });
  builder.papxAll(text);
  builder.ccp.txbx = boxText.length;
  const cps: number[] = [];
  const spas: Uint8Array[] = [];
  const records: Uint8Array[] = [];
  let from = -1;
  let run = -1;
  for (let cp = 0; cp <= text.length; cp++) {
    const drawing = cp < text.length && text.charCodeAt(cp) === 8;
    if (drawing) {
      const id = 5000 + cps.length;
      records.push(boxShape(id, shapes[cps.length]));
      cps.push(cp);
      spas.push(spa(id));
      if (from < 0) from = cp;
    } else if (from >= 0) {
      builder.chpx(from, cp, sprm(0x0855, 1));
      from = -1;
    }
    run = cp;
  }
  assert.equal(cps.length, shapes.length);
  cps.push(run);
  builder.table(FC.PlcSpaMom, plc(cps, spas));
  builder.table(FC.DggInfo, cat(art(0xf, 0, 0xf000), u8s(0), art(0xf, 0, 0xf002, art(0xf, 0, 0xf003, art(0xf, 0, 0xf004, art(2, 75, 0xf00a, u32le(4999), u32le(0x05))), ...records))));
  const starts = [0];
  for (const box of boxes) starts.push(starts[starts.length - 1] + box.length);
  starts.push(boxText.length);
  builder.table(FC.PlcftxbxTxt, plc(starts, starts.slice(1).map(() => new Uint8Array(22))));
  return builder;
}

test('the text of a box goes to the first shape that names it: other shapes get none, at any depth', () => {
  // Two shapes in the body show the first box. That box holds 2,500 shapes which all show the second box.
  const inner = 2500;
  let second = '';
  for (let i = 0; i < 200; i++) second += 'p' + String(i).padStart(3, '0') + ' ' + 'x'.repeat(40) + '\r';
  const shapes = [1, 1];
  for (let i = 0; i < inner; i++) shapes.push(2);
  const builder = boxDoc('A\u0008\u0008\r', ['\u0008'.repeat(inner) + '\r', second], shapes);
  const started = Date.now();
  const model = builder.read();
  const elapsed = Date.now() - started;
  assert.deepEqual(textOf(model.body.blocks), ['A']);
  const [first, other] = drawingsOf(paragraphs(model.body.blocks)[0]);
  assert.equal(first.textBoxes!.length, 1);
  assert.equal(other.textBoxes, undefined);
  // In the box: one of the shapes has the text of the second box, the others have none.
  const holders = drawingsOf(paragraphs(first.textBoxes![0])[0]);
  assert.equal(holders.length, inner);
  const withText = holders.filter((drawing) => drawing.textBoxes !== undefined);
  assert.equal(withText.length, 1);
  assert.equal(withText[0], holders[0]);
  assert.equal(textOf(withText[0].textBoxes![0]).length, 200);
  assert.equal(textOf(withText[0].textBoxes![0])[199].slice(0, 4), 'p199');
  // Reading the second box again for every shape took seconds.
  assert.ok(elapsed < 2500, `took ${elapsed} ms`);
});

test('a text box whose text holds a shape that shows the same text ends there', () => {
  const model = boxDoc('A\u0008\r', ['in\u0008side\r'], [1, 1]).read();
  const [outer] = drawingsOf(paragraphs(model.body.blocks)[0]);
  assert.deepEqual(textOf(outer.textBoxes![0]), ['inside']);
  const [nested] = drawingsOf(paragraphs(outer.textBoxes![0])[0]);
  assert.equal(nested.textBoxes, undefined);
});

// ---------------------------------------------------------------------------
// Pictures

/** A PICF with what follows it: the natural size in twips, the scale in thousandths, the cropping in twips. */
function picf(width: number, height: number, scale: [number, number], crop: [number, number, number, number], ...records: Uint8Array[]): Uint8Array {
  const header = new Uint8Array(68);
  const view = new DataView(header.buffer);
  const body = cat(...records);
  view.setUint32(0, 68 + body.length, true);
  view.setUint16(4, 68, true);
  view.setUint16(6, 0x64, true);
  view.setUint16(28, width, true);
  view.setUint16(30, height, true);
  view.setUint16(32, scale[0], true);
  view.setUint16(34, scale[1], true);
  crop.forEach((value, i) => view.setInt16(36 + i * 2, value, true));
  return cat(header, body);
}

const be32 = (value: number): number[] => [(value >>> 24) & 0xff, (value >>> 16) & 0xff, (value >>> 8) & 0xff, value & 0xff];
const be16 = (value: number): number[] => [(value >> 8) & 0xff, value & 0xff];

/** The start of a PNG file of the given size in pixels. */
const png = (width: number, height: number): Uint8Array => cat([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], be32(13), [0x49, 0x48, 0x44, 0x52], be32(width), be32(height), [8, 2, 0, 0, 0], new Uint8Array(8));

/** The start of a JPEG file: an application segment, a table, then the frame header `marker` with the size. */
const jpeg = (width: number, height: number, marker = 0xc0): Uint8Array =>
  cat([0xff, 0xd8], [0xff, 0xe0], be16(16), new Uint8Array(14), [0xff, 0xff, 0xff, 0xdb], be16(67), new Uint8Array(65), [0xff, marker], be16(11), [8], be16(height), be16(width), new Uint8Array(6), [0xff, 0xda]);

const bitmapBlip = (type: number, instance: number, data: Uint8Array): Uint8Array => art(0, instance, type, new Uint8Array(16), u8s(0xff), data);
const pictureShape = art(0xf, 0, 0xf004, art(2, 75, 0xf00a, u32le(1024), u32le(0xa00)), art(3, 1, 0xf00b, cat(u16le(0x0104), u32le(1))));

function inlinePicture(data: Uint8Array): DrawingItem {
  const builder = new DocBuilder();
  const text = 'A\u0001B\r';
  builder.text(text);
  builder.papxAll(text);
  builder.dataBlob(new Uint8Array(21));
  const at = builder.dataBlob(data);
  builder.chpx(1, 2, sprm(0x0855, 1), sprm(0x6a03, u32le(at)));
  const [drawing] = drawingsOf(paragraphs(builder.read().body.blocks)[0]);
  return drawing;
}

const TWIP = 635;

test('a picture scaled beyond 22 inches on one axis keeps the shape its own header states, not that of the clipped natural size', () => {
  // 2369 x 1449 pixels. Its natural width would be 35535 twips, which the file cuts to 31680; the height is 21778.
  // Shown at a quarter of its height, with a width scale that gives metres.
  const height = (21778 * 250) / 1000;
  const expected = { width: Math.round(((height * 2369) / 1449) * TWIP), height: Math.round(height * TWIP) };
  const fromPng = inlinePicture(picf(31680, 21778, [3000, 250], [0, 0, 0, 0], pictureShape, bitmapBlip(0xf01e, 0x6e0, png(2369, 1449))));
  assert.deepEqual({ width: fromPng.width, height: fromPng.height }, expected);
  // About 8902 x 5445 twips: the shape of the picture. The natural size would have given 7920.
  assert.ok(Math.abs(fromPng.width! / TWIP - 8902) <= 1, String(fromPng.width! / TWIP));
  assert.equal(fromPng.image!.target, '/doc/media/image1.png');

  for (const marker of [0xc0, 0xc1, 0xc2, 0xc3]) {
    const fromJpeg = inlinePicture(picf(31680, 21778, [3000, 250], [0, 0, 0, 0], pictureShape, bitmapBlip(0xf01d, 0x46a, jpeg(2369, 1449, marker))));
    assert.deepEqual({ width: fromJpeg.width, height: fromJpeg.height }, expected, `marker ${marker.toString(16)}`);
  }

  // The other axis: the width is what the scale gives, the height follows the picture.
  const tall = inlinePicture(picf(20000, 31680, [100, 4000], [0, 0, 0, 0], pictureShape, bitmapBlip(0xf01e, 0x6e0, png(1000, 3000))));
  assert.deepEqual({ width: tall.width, height: tall.height }, { width: 2000 * TWIP, height: 6000 * TWIP });

  // Cropping takes its share from each side of the picture.
  const cropped = inlinePicture(picf(30000, 20000, [3000, 500], [15000, 0, 0, 0], pictureShape, bitmapBlip(0xf01e, 0x6e0, png(3000, 1000))));
  // Half the width is cut off: 1500 x 1000 pixels are shown, 10000 twips high.
  assert.deepEqual({ width: cropped.width, height: cropped.height }, { width: 15000 * TWIP, height: 10000 * TWIP });

  // A picture that does not say how large it is: the natural size of the file gives the shape, as before.
  const silent = inlinePicture(picf(31680, 21778, [3000, 250], [0, 0, 0, 0], pictureShape, bitmapBlip(0xf01e, 0x6e0, cat([0x89, 0x50, 0x4e, 0x47], new Uint8Array(40)))));
  assert.equal(silent.width, Math.round(((height * 31680) / 21778) * TWIP));
  const noFrame = inlinePicture(picf(31680, 21778, [3000, 250], [0, 0, 0, 0], pictureShape, bitmapBlip(0xf01d, 0x46a, cat([0xff, 0xd8, 0xff, 0xe0], be16(16), new Uint8Array(14), [0xff, 0xda], new Uint8Array(20)))));
  assert.equal(noFrame.width, Math.round(((height * 31680) / 21778) * TWIP));

  // Both axes beyond the limit: 22 inches each way, whatever the picture says.
  const both = inlinePicture(picf(31680, 21778, [3000, 3000], [0, 0, 0, 0], pictureShape, bitmapBlip(0xf01e, 0x6e0, png(2369, 1449))));
  assert.deepEqual({ width: both.width, height: both.height }, { width: 31680 * TWIP, height: 31680 * TWIP });
  // Within the limit nothing is changed: the scales of the file hold, even where they distort the picture.
  const within = inlinePicture(picf(10000, 5000, [500, 1000], [0, 0, 0, 0], pictureShape, bitmapBlip(0xf01e, 0x6e0, png(100, 100))));
  assert.deepEqual({ width: within.width, height: within.height }, { width: 5000 * TWIP, height: 5000 * TWIP });
});

test('pictureSize: the size a picture states in its own header, for every kind the reader hands out', async () => {
  const { pictureSize } = await import('../src/formats/doc/drawings');
  assert.deepEqual(pictureSize({ type: 'image/png', bytes: png(640, 480) }), { width: 640, height: 480 });
  assert.equal(pictureSize({ type: 'image/png', bytes: png(640, 480).subarray(0, 23) }), undefined);
  assert.equal(pictureSize({ type: 'image/png', bytes: png(0, 480) }), undefined);
  assert.deepEqual(pictureSize({ type: 'image/jpeg', bytes: jpeg(800, 600) }), { width: 800, height: 600 });
  assert.deepEqual(pictureSize({ type: 'image/jpeg', bytes: jpeg(800, 600, 0xc2) }), { width: 800, height: 600 });
  // A table of Huffman codes (0xC4) is not a frame header.
  assert.equal(pictureSize({ type: 'image/jpeg', bytes: jpeg(800, 600, 0xc4) }), undefined);
  // Cut off inside the frame header, and a segment that claims a length of nothing.
  assert.equal(pictureSize({ type: 'image/jpeg', bytes: jpeg(800, 600).subarray(0, 96) }), undefined);
  assert.equal(pictureSize({ type: 'image/jpeg', bytes: cat([0xff, 0xd8, 0xff, 0xe0, 0, 0], new Uint8Array(40)) }), undefined);
  // Thousands of segments before the frame header: the search gives up; it does not run through all of them.
  const many: number[] = [0xff, 0xd8];
  for (let i = 0; i < 5000; i++) many.push(0xff, 0xe1, 0, 2);
  assert.equal(pictureSize({ type: 'image/jpeg', bytes: cat(many, jpeg(800, 600).subarray(2)) }), undefined);
  assert.deepEqual(pictureSize({ type: 'image/jpeg', bytes: cat(many.slice(0, 2 + 4 * 100), jpeg(800, 600).subarray(2)) }), { width: 800, height: 600 });

  // BMP: the file header, then a header with 32-bit sizes (the height negative for rows from the top) or 16-bit ones.
  const bmp = new Uint8Array(54);
  const view = new DataView(bmp.buffer);
  bmp.set([0x42, 0x4d]);
  view.setUint32(14, 40, true);
  view.setInt32(18, 320, true);
  view.setInt32(22, -200, true);
  assert.deepEqual(pictureSize({ type: 'image/bmp', bytes: bmp }), { width: 320, height: 200 });
  view.setUint32(14, 12, true);
  view.setUint16(18, 64, true);
  view.setUint16(20, 48, true);
  assert.deepEqual(pictureSize({ type: 'image/bmp', bytes: bmp }), { width: 64, height: 48 });

  // EMF: the frame of the header record. WMF: the box of a placeable metafile.
  const emf = new Uint8Array(88);
  const emfView = new DataView(emf.buffer);
  emfView.setUint32(0, 1, true);
  emfView.setInt32(24, 100, true);
  emfView.setInt32(28, 200, true);
  emfView.setInt32(32, 2100, true);
  emfView.setInt32(36, 1200, true);
  emf.set([0x20, 0x45, 0x4d, 0x46], 40);
  assert.deepEqual(pictureSize({ type: 'image/x-emf', bytes: emf }), { width: 2000, height: 1000 });
  emf[40] = 0;
  assert.equal(pictureSize({ type: 'image/x-emf', bytes: emf }), undefined);
  const wmf = new Uint8Array(22);
  const wmfView = new DataView(wmf.buffer);
  wmfView.setUint32(0, 0x9ac6cdd7, true);
  wmfView.setInt16(6, -10, true);
  wmfView.setInt16(8, 0, true);
  wmfView.setInt16(10, 590, true);
  wmfView.setInt16(12, 300, true);
  assert.deepEqual(pictureSize({ type: 'image/x-wmf', bytes: wmf }), { width: 600, height: 300 });
  assert.equal(pictureSize({ type: 'image/x-wmf', bytes: new Uint8Array(22) }), undefined);
  assert.equal(pictureSize({ type: 'image/tiff', bytes: new Uint8Array(64) }), undefined);
});

// ---------------------------------------------------------------------------
// Toggles of styles

const BOLD = 0x0835;
const ITALIC = 0x0836;
const CHAR_STYLE = 0x4a30;

/** Styles: 0 Normal, 1 a bold paragraph style, then character styles: 2 bold, 3 not bold, 4 the opposite, 5 as below, 6 a table style with bold text, 7 based on 2 and italic. */
const TOGGLE_STYLES = stsh([
  { sti: 0, name: 'Normal' },
  { name: 'BoldPara', base: 0, chpx: sprm(BOLD, 1) },
  { kind: 2, name: 'BoldChar', chpx: sprm(BOLD, 1) },
  { kind: 2, name: 'OffChar', chpx: sprm(BOLD, 0) },
  { kind: 2, name: 'FlipChar', chpx: sprm(BOLD, 0x81) },
  { kind: 2, name: 'SameChar', chpx: cat(sprm(BOLD, 0x80), sprm(ITALIC, 1)) },
  { kind: 3, name: 'BoldTable', chpx: sprm(BOLD, 1) },
  { kind: 2, name: 'Derived', base: 2, chpx: sprm(ITALIC, 1) },
]);

/** One paragraph of the style `istd` with a run for each list of sprms; returns the bold value in effect for each run. */
function boldOf(istd: number, runs: Uint8Array[][]): { model: DocModel; effective: boolean[]; direct: (boolean | undefined)[] } {
  const builder = new DocBuilder();
  const text = runs.map((_, i) => String.fromCharCode(97 + i)).join('') + '\r';
  builder.text(text);
  builder.papx(0, text.length, istd);
  runs.forEach((sprms, i) => builder.chpx(i, i + 1, ...sprms));
  builder.table(FC.Stshf, TOGGLE_STYLES);
  const model = builder.read();
  const paragraph = paragraphs(model.body.blocks)[0];
  const all = runsOf(paragraph);
  assert.equal(all.length, runs.length);
  return {
    model,
    effective: all.map((run) => resolveRunProps(model, paragraph.props, run.props).bold === true),
    direct: all.map((run) => run.props.bold),
  };
}

test('a toggle of a character style is on or off, whatever the paragraph style says', () => {
  const style = (istd: number) => sprm(CHAR_STYLE, u16le(istd));
  // In a bold paragraph: a bold character style stays bold, one that says "not bold" is not, "the opposite" is not, "as below" is.
  const inBold = boldOf(1, [[], [style(2)], [style(3)], [style(4)], [style(5)], [style(7)]]);
  assert.deepEqual(inBold.effective, [true, true, false, false, true, true]);
  // In a plain paragraph: bold, not bold, the opposite of not bold, as below.
  const inPlain = boldOf(0, [[], [style(2)], [style(3)], [style(4)], [style(5)], [style(7)]]);
  assert.deepEqual(inPlain.effective, [false, true, false, true, false, true]);
  // Where the rule of the model gives the same, the run states nothing itself.
  assert.deepEqual(inPlain.direct, [undefined, undefined, undefined, undefined, undefined, undefined]);
  assert.deepEqual(inBold.direct, [undefined, true, false, undefined, undefined, true]);
  // The styles of the model: "the opposite" is what `true` means in a character style there.
  const styles = inBold.model.styles.styles;
  assert.deepEqual(styles.get('BoldChar')!.run, { bold: true });
  assert.deepEqual(styles.get('OffChar')!.run, { bold: false });
  assert.deepEqual(styles.get('FlipChar')!.run, { bold: true });
  assert.deepEqual(styles.get('SameChar')!.run, { italic: true });
  // The other properties of the run are untouched, and runs that say the same still share their properties.
  const twice = boldOf(1, [[style(2), sprm(0x4a43, u16le(30))], [], [style(2), sprm(0x4a43, u16le(30))]]);
  const runs = runsOf(paragraphs(twice.model.body.blocks)[0]);
  assert.deepEqual(runs[0].props, { style: 'BoldChar', size: 30, bold: true });
  assert.equal(runs[0].props, runs[2].props);
});

test('a toggle a run states as "the style" or "the opposite of the style" follows the styles as this format reads them', () => {
  const style = (istd: number) => sprm(CHAR_STYLE, u16le(istd));
  const opposite = sprm(BOLD, 0x81);
  const same = sprm(BOLD, 0x80);
  // Bold paragraph: the opposite of bold; of bold over bold; of "not bold" over bold; of "the opposite" over bold.
  assert.deepEqual(boldOf(1, [[opposite], [style(2), opposite], [style(3), opposite], [style(4), opposite]]).effective, [false, false, true, true]);
  assert.deepEqual(boldOf(0, [[opposite], [style(2), opposite], [style(3), opposite], [style(4), opposite]]).effective, [true, false, true, false]);
  // "As the style": the value the styles give.
  assert.deepEqual(boldOf(1, [[same], [style(2), same], [style(3), same], [style(4), same]]).effective, [true, true, false, false]);
  // On and off in the run itself hold against every style; a later sprm replaces an earlier one.
  assert.deepEqual(boldOf(1, [[sprm(BOLD, 0)], [style(2), sprm(BOLD, 0)], [style(3), sprm(BOLD, 1)], [opposite, sprm(BOLD, 1)], [sprm(BOLD, 1), opposite]]).effective, [false, false, true, true, false]);
});

test('toggles of a table style count as the lowest level: a paragraph or character style on top is not flipped by it', () => {
  const builder = new DocBuilder();
  // One row of two cells: "ab" in a plain paragraph, "cdef" in a bold one.
  const text = 'ab\u0007cdef\u0007\u0007x\r';
  builder.text(text);
  builder.papx(0, 3, 0, IN_TABLE).papx(3, 8, 1, IN_TABLE);
  builder.papx(8, 9, 0, IN_TABLE, ROW_END, defTable([0, 2000, 4000]), sprm(0x563a, u16le(6)));
  builder.papx(9, 11, 0);
  const style = (istd: number) => sprm(CHAR_STYLE, u16le(istd));
  builder.chpx(1, 2, sprm(BOLD, 0x81));
  builder.chpx(4, 5, style(2)).chpx(5, 6, style(3)).chpx(6, 7, sprm(BOLD, 0x81));
  // The mark of the second cell is "the opposite of the style" too.
  builder.chpx(7, 8, sprm(BOLD, 0x81));
  builder.table(FC.Stshf, TOGGLE_STYLES);
  const model = builder.read();
  const table = only(model);
  assert.equal(table.props.style, 'BoldTable');
  const bold = (cell: number) => {
    const paragraph = paragraphs(table.rows[0].cells[cell].blocks)[0];
    return runsOf(paragraph).map((run) => resolveRunProps(model, paragraph.props, run.props, { table, row: 0, cell }).bold === true);
  };
  // Plain paragraph in the bold table: bold; the opposite of that: not bold.
  assert.deepEqual(bold(0), [true, false]);
  // Bold paragraph in the bold table: still bold (the model alone would flip it); bold character style: bold;
  // "not bold": not bold; the opposite of bold: not bold.
  assert.deepEqual(bold(1), [true, true, false, false]);
  const second = paragraphs(table.rows[0].cells[1].blocks)[0];
  assert.equal(resolveRunProps(model, second.props, second.props.markProps!, { table, row: 0, cell: 1 }).bold, false);
  // Outside the table nothing of this applies.
  const after = paragraphs(model.body.blocks)[0];
  assert.equal(resolveRunProps(model, after.props, runsOf(after)[0].props).bold, undefined);
});

// ---------------------------------------------------------------------------
// Smaller things

test('number formats: every value of the list has the name a docx file gives it', () => {
  const expected: [number, string][] = [
    [0, 'decimal'],
    [1, 'upperRoman'],
    [2, 'lowerRoman'],
    [3, 'upperLetter'],
    [4, 'lowerLetter'],
    [5, 'ordinal'],
    [6, 'cardinalText'],
    [7, 'ordinalText'],
    [8, 'hex'],
    [9, 'chicago'],
    [10, 'ideographDigital'],
    [11, 'japaneseCounting'],
    [12, 'aiueo'],
    [13, 'iroha'],
    [14, 'decimalFullWidth'],
    [15, 'decimalHalfWidth'],
    [18, 'decimalEnclosedCircle'],
    [22, 'decimalZero'],
    [23, 'bullet'],
    [24, 'ganada'],
    [25, 'chosung'],
    [30, 'ideographTraditional'],
    [31, 'ideographZodiac'],
    [37, 'chineseCounting'],
    [39, 'chineseCountingThousand'],
    [41, 'koreanDigital'],
    [45, 'hebrew1'],
    [46, 'arabicAlpha'],
    [47, 'hebrew2'],
    [48, 'arabicAbjad'],
    [53, 'thaiLetters'],
    [56, 'vietnameseCounting'],
    [57, 'numberInDash'],
    [58, 'russianLower'],
    [59, 'russianUpper'],
    [255, 'none'],
    // Values without a name count in decimal numbers.
    [40, 'decimal'],
    [60, 'decimal'],
    [200, 'decimal'],
  ];
  for (const [nfc, name] of expected) assert.equal(numberFormat(nfc), name, String(nfc));
});

test('alignment: stated as a side of the page it changes sides in a right-to-left paragraph; stated as start and end it does not', () => {
  const builder = new DocBuilder();
  const text = 'a\rb\rc\rd\re\rf\rg\rh\r';
  builder.text(text);
  const physical = (value: number) => sprm(0x2403, value);
  const logical = (value: number) => sprm(0x2461, value);
  const bidi = sprm(0x2441, 1);
  builder.table(FC.Stshf, stsh([{ sti: 0, name: 'Normal' }, { name: 'Rtl', base: 0, papx: bidi }]));
  builder.papx(0, 2, 0, physical(2)); // right, left to right
  builder.papx(2, 4, 0, bidi, physical(2)); // the right side of a right-to-left paragraph: where it starts
  builder.papx(4, 6, 0, physical(0), bidi); // the left side: where it ends, whatever the order of the sprms
  builder.papx(6, 8, 0, bidi, physical(2), logical(2)); // both stated: the second holds
  builder.papx(8, 10, 0, bidi, logical(0), physical(2));
  builder.papx(10, 12, 0, bidi, physical(1)); // centred is centred
  builder.papx(12, 14, 1, physical(2)); // right to left by its style
  builder.papx(14, 16, 1, sprm(0x2441, 0), physical(2)); // the style undone
  const model = builder.read();
  assert.deepEqual(
    paragraphs(model.body.blocks).map((paragraph) => paragraph.props.align),
    ['right', 'left', 'right', 'right', 'left', 'center', 'left', 'right'],
  );
});

test('a section starts its page numbers anew only when it says so', () => {
  const sed = (builder: DocBuilder, ...sprms: Uint8Array[]): Uint8Array => {
    const list = cat(...sprms);
    const at = builder.blob(cat(u16le(list.length), list));
    return cat(u16le(0), u32le(at), u16le(0), u32le(0xffffffff));
  };
  const read = (...sprms: Uint8Array[]) => {
    const builder = new DocBuilder();
    builder.text('x\r');
    builder.papx(0, 2, 0);
    builder.table(FC.PlcfSed, plc([0, 2], [sed(builder, ...sprms)]));
    return builder.read().body.section!.pageNumbering;
  };
  // A first page number without the flag is what is left from an earlier state of the section: it does not apply.
  assert.equal(read(sprm(0x501c, u16le(5))), undefined);
  assert.equal(read(sprm(0x7044, u32le(7))), undefined);
  assert.equal(read(sprm(0x3011, 0), sprm(0x501c, u16le(5))), undefined);
  assert.deepEqual(read(sprm(0x300e, 2), sprm(0x501c, u16le(5))), { format: 'lowerRoman' });
  // With the flag it does, in either order and in both sizes.
  assert.deepEqual(read(sprm(0x3011, 1), sprm(0x501c, u16le(5))), { start: 5 });
  assert.deepEqual(read(sprm(0x501c, u16le(5)), sprm(0x3011, 1)), { start: 5 });
  assert.deepEqual(read(sprm(0x3011, 1), sprm(0x7044, u32le(70000))), { start: 70000 });
  // The flag alone restarts nothing the model could state.
  assert.equal(read(sprm(0x3011, 1)), undefined);
});

test('the indent of a table is that of its leftmost edge: a first row that starts further right does not move it', () => {
  // The first row is one cell far to the right (it has grid columns before it) and states the indent of the table;
  // the rows after it start at the left edge.
  const rows = [['top'], ['a', 'b'], ['c', 'd']];
  const table = only(
    tableDoc(rows, (row) => (row === 0 ? [dxaLeft(8532), gapHalf(108), defTable([8424, 10500]), widthIndent(-223)] : [dxaLeft(-223), gapHalf(108), defTable([-331, 5000, 10500]), widthIndent(-223)])).read(),
  );
  assert.deepEqual(table.props.indent, { type: 'dxa', value: -223 });
  assert.deepEqual(table.grid, [5331, 3424, 2076]);
  assert.equal(table.rows[0].props.gridBefore, 2);
  assert.deepEqual(table.rows[0].props.widthBefore, { type: 'dxa', value: 8755 });
  assert.equal(table.rows[1].props.gridBefore, undefined);

  // The rows at the left edge state no indent: the place of their first text is the indent.
  const old = only(tableDoc(rows, (row) => (row === 0 ? [dxaLeft(8532), gapHalf(108), defTable([8424, 10500]), widthIndent(-223)] : [dxaLeft(-223), gapHalf(108), defTable([-331, 5000, 10500])])).read());
  assert.deepEqual(old.props.indent, { type: 'dxa', value: -223 });

  // No row at the left edge says where its text starts: the edge plus the cell margin in effect.
  const bare = only(tableDoc(rows, (row) => (row === 0 ? [defTable([8424, 10500]), widthIndent(-223)] : [gapHalf(70), defTable([-331, 5000, 10500])])).read());
  assert.deepEqual(bare.props.indent, { type: 'dxa', value: -331 + 70 });

  // The first row at the left edge still decides, whatever the rows after it state.
  const first = only(tableDoc(rows, (row) => (row === 0 ? [dxaLeft(93), gapHalf(108), defTable([-15, 9000]), widthIndent(93)] : [dxaLeft(500), gapHalf(108), defTable([392, 5000, 9000]), widthIndent(700)])).read());
  assert.deepEqual(first.props.indent, { type: 'dxa', value: 93 });
});
