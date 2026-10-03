// Shapes of doc and RTF files: their geometry, outline and fill in the model,
// on the page, on the canvas, and in the docx they are converted to.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { convertForEditing } from '../src/formats/convert';
import { open } from '../src/index';
import { DocLayout } from '../src/layout/doc/layout';
import type { ImageItem, LayoutPage, PageItem, ShapeItem } from '../src/layout/doc/types';
import { officeArtColor, officeArtLook } from '../src/model/doc/shapes';
import type { DocModel, DrawingItem, Paragraph, ShapeLook } from '../src/model/doc/types';
import { renderPage } from '../src/render/doc';
import { OpcPackage } from '../src/opc';
import { DocBuilder, FC, cat, plc, sprm, u16le, u32le, u8s } from './doc-binary-fixtures';
import { fakeMeasurer, lay, lines, p, rects, texts } from './doc-layout-fixtures';
import { bytesOf } from './rtf-fixtures';

// ------------------------------------------------------------------ fixtures

/** An OfficeArt record: version, instance, type and content. */
const art = (version: number, instance: number, type: number, ...content: Uint8Array[]): Uint8Array => {
  const body = cat(...content);
  return cat(u16le(version | (instance << 4)), u16le(type), u32le(body.length), body);
};

const options = (properties: [number, number][]): Uint8Array => art(3, properties.length, 0xf00b, ...properties.map(([id, value]) => cat(u16le(id), u32le(value))));

interface DocShape {
  /** MSOSPT. */
  type: number;
  /** Flags of the shape: 0x40 flipped horizontally, 0x80 vertically. */
  flags?: number;
  props?: [number, number][];
  /** Left, top, right, bottom in twips, from the column and the paragraph. */
  rect: [number, number, number, number];
  /** The shape is a text box with this text. */
  text?: string;
  /** Replaces the property table. */
  table?: Uint8Array;
}

/** A doc file whose only paragraph anchors the shapes. */
function shapesDoc(shapes: DocShape[]): DocBuilder {
  const builder = new DocBuilder();
  const body = 'Anchor' + '\u0008'.repeat(shapes.length) + '\r';
  const boxed = shapes.filter((shape) => shape.text !== undefined);
  const boxes = boxed.map((shape) => shape.text + '\r');
  const boxText = boxes.length ? boxes.join('') + '\r' : '';
  const text = body + boxText + '\r';
  builder.text(text);
  builder.papxAll(text);
  builder.chpx(6, 6 + shapes.length, sprm(0x0855, 1));
  if (boxText) builder.ccp.txbx = boxText.length;
  const id = (i: number) => 1025 + i;
  // Relative to the column and the paragraph, no wrapping.
  const spa = (i: number) => cat(u32le(id(i)), ...shapes[i].rect.map(u32le), u16le((2 << 1) | (2 << 3) | (3 << 5)), u32le(0));
  builder.table(FC.PlcSpaMom, plc([...shapes.map((_, i) => 6 + i), body.length], shapes.map((_, i) => spa(i))));
  const containers = shapes.map((shape, i) => {
    const props: [number, number][] = shape.text !== undefined ? [[0x0080, (boxed.indexOf(shape) + 1) << 16], ...(shape.props || [])] : shape.props || [];
    return art(0xf, 0, 0xf004, art(2, shape.type, 0xf00a, u32le(id(i)), u32le(0xa00 | (shape.flags || 0))), shape.table || options(props));
  });
  const patriarch = art(0xf, 0, 0xf004, art(2, 0, 0xf00a, u32le(1024), u32le(0x05)));
  builder.table(FC.DggInfo, cat(art(0xf, 0, 0xf000, art(0, 0, 0xf006, new Uint8Array(16))), u8s(0), art(0xf, 0, 0xf002, art(0, 1, 0xf008, new Uint8Array(8)), art(0xf, 0, 0xf003, patriarch, ...containers))));
  if (boxed.length) {
    const entry = (shapeId: number) => {
      const bytes = new Uint8Array(22);
      bytes.set(u32le(shapeId), 14);
      return bytes;
    };
    const cps = [0];
    for (const box of boxes) cps.push(cps[cps.length - 1] + box.length);
    cps.push(boxText.length);
    builder.table(FC.PlcftxbxTxt, plc(cps, [...boxed.map((shape) => entry(id(shapes.indexOf(shape)))), entry(0)]));
  }
  return builder;
}

const P_ROTATION = 0x0004;
const P_ADJUST = 0x0147;
const P_FILL_TYPE = 0x0180;
const P_FILL_COLOR = 0x0181;
const P_FILL_OPACITY = 0x0182;
const P_PICTURE = 0x0104;
const P_FILL_FLAGS = 0x01bf;
const P_LINE_COLOR = 0x01c0;
const P_LINE_WIDTH = 0x01cb;
const P_LINE_DASHING = 0x01ce;
const P_LINE_FLAGS = 0x01ff;
/** The flag and the bit that says it is stated. */
const FILLED = 0x100010;
const NOT_FILLED = 0x100000;
const LINED = 0x080008;
const NOT_LINED = 0x080000;

const sp = (name: string, value: string | number) => `{\\sp{\\sn ${name}}{\\sv ${value}}}`;
/** An RTF shape: its rectangle in twips from the column and the paragraph, its properties and the text of its box. */
const shp = (rect: number[], props: string, text?: string) =>
  `{\\shp{\\*\\shpinst\\shpleft${rect[0]}\\shptop${rect[1]}\\shpright${rect[2]}\\shpbottom${rect[3]}\\shpbxcolumn\\shpbypara\\shpwr3${props}${text === undefined ? '' : `{\\shptxt ${text}\\par}`}}}`;
/** A page of 400 x 400 px with margins of 40 px, as the layout fixtures have it. */
const rtf = (body: string) => bytesOf(`{\\rtf1\\ansi\\paperw6000\\paperh6000\\margl600\\margr600\\margt600\\margb600 ${body}\\par}`);

function drawingsIn(model: DocModel): DrawingItem[] {
  const out: DrawingItem[] = [];
  for (const block of model.body.blocks) {
    if (block.type !== 'paragraph') continue;
    for (const inline of (block as Paragraph).children) if (inline.type === 'run') for (const item of inline.content) if (item.type === 'drawing') out.push(item);
  }
  return out;
}

const looks = (model: DocModel): (ShapeLook | undefined)[] => drawingsIn(model).map((item) => item.shape);

async function rtfModel(body: string): Promise<DocModel> {
  return (await open(rtf(body))).document!();
}

function pageOf(model: DocModel): LayoutPage {
  const layout = new DocLayout(model, fakeMeasurer);
  assert.equal(layout.layout().done, true);
  return layout.pages[0];
}

const shapesOf = (page: LayoutPage): ShapeItem[] => page.items.filter((item): item is ShapeItem => item.type === 'shape');
const placeholders = (page: LayoutPage): ImageItem[] => page.items.filter((item): item is ImageItem => item.type === 'image' && item.placeholder !== undefined);

/** What a page paints, without the marks of its lines and what only selection needs. */
const painted = (page: LayoutPage): string[] =>
  page.items.filter((item) => item.type !== 'mark').map((item) => JSON.stringify(item, (key, value) => (key === 'source' ? undefined : typeof value === 'number' ? Math.round(value * 100) / 100 : value)));

// ------------------------------------------------------- OfficeArt properties

test('shapes: a shape that states nothing is a box filled in white and outlined in black, 0.75 pt wide', () => {
  assert.deepEqual(officeArtLook({ type: 1 }), { geometry: 'rect', fill: 'FFFFFF', line: { color: '000000', width: 9525 } });
  // A file that does not say what the shape is means a box; a text box and a picture frame are boxes too.
  for (const type of [undefined, 202, 75]) assert.equal(officeArtLook({ type }).geometry, 'rect');
  assert.equal(officeArtLook({ type: 2 }).geometry, 'roundRect');
  assert.equal(officeArtLook({ type: 3 }).geometry, 'ellipse');
  assert.equal(officeArtLook({ type: 20 }).geometry, 'line');
  assert.equal(officeArtLook({ type: 32 }).geometry, 'line');
  // Everything else (a freeform, an arrow, a star) is a geometry the reader does not know.
  for (const type of [0, 4, 13, 187, 100000, -1]) assert.equal(officeArtLook({ type }).geometry, 'other');
});

test('shapes: fill and outline are taken as stated; a line has no fill', () => {
  assert.deepEqual(officeArtLook({ type: 1, fFilled: false, fLine: false }), { geometry: 'rect' });
  assert.deepEqual(officeArtLook({ type: 3, fillColor: 0x0000ff, lineColor: 0xff8000, lineWidth: 25400, lineDashing: 6 }), {
    geometry: 'ellipse',
    fill: 'FF0000',
    line: { color: '0080FF', width: 25400, dash: 'dash' },
  });
  assert.deepEqual(officeArtLook({ type: 20, fFilled: true, fillColor: 0x00ff00 }), { geometry: 'line', line: { color: '000000', width: 9525 } });
  // A width of nothing is the thinnest line; one that is not a width is the default; one beyond what Word draws is cut.
  assert.equal(officeArtLook({ type: 1, lineWidth: 0 }).line!.width, 0);
  assert.equal(officeArtLook({ type: 1, lineWidth: -5 }).line!.width, 9525);
  assert.equal(officeArtLook({ type: 1, lineWidth: NaN }).line!.width, 9525);
  assert.equal(officeArtLook({ type: 1, lineWidth: 0xffffffff }).line!.width, 20116800);
  // The dashes by their numbers; a number that is none is a solid line.
  const dashes = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, -1, 1000].map((lineDashing) => officeArtLook({ type: 1, lineDashing }).line!.dash);
  assert.deepEqual(dashes, [undefined, 'sysDash', 'sysDot', 'sysDashDot', 'sysDashDotDot', 'dot', 'dash', 'lgDash', 'dashDot', 'lgDashDot', 'lgDashDotDot', undefined, undefined, undefined]);
  // The corners of a rounded rectangle, from 21600ths to 100000ths of the shorter side.
  assert.equal(officeArtLook({ type: 2 }).corner, undefined);
  assert.equal(officeArtLook({ type: 2, adjust: 5400 }).corner, 25000);
  assert.equal(officeArtLook({ type: 2, adjust: 50000 }).corner, undefined);
  assert.equal(officeArtLook({ type: 1, adjust: 5400 }).corner, undefined);
});

test('shapes: a colour that is an index or a system colour is not taken for a colour', () => {
  assert.equal(officeArtColor(0x00a1b2c3), 'C3B2A1');
  assert.equal(officeArtColor(0x02a1b2c3), 'C3B2A1');
  for (const value of [0x01000005, 0x08000001, 0x100000f0, 0x04a1b2c3, undefined, NaN]) assert.equal(officeArtColor(value), undefined);
  // Such a colour is not made up: the fill is left out, the outline is black, and the look says what it lost.
  assert.deepEqual(officeArtLook({ type: 1, fillColor: 0x08000009, lineColor: 0x100000f4 }), { geometry: 'rect', line: { color: '000000', width: 9525 }, lost: ['color'] });
  assert.deepEqual(officeArtLook({ type: 1, fillColor: 0x01000003 }), { geometry: 'rect', line: { color: '000000', width: 9525 }, lost: ['color'] });
  assert.deepEqual(officeArtLook({ type: 1, lineColor: 0x08000001 }), { geometry: 'rect', fill: 'FFFFFF', line: { color: '000000', width: 9525 }, lost: ['color'] });
  // A colour of a fill or an outline that is switched off is not looked at.
  assert.deepEqual(officeArtLook({ type: 1, fFilled: false, fLine: false, fillColor: 0x08000009, lineColor: 0x08000001 }), { geometry: 'rect' });
});

test('shapes: a fill that is not a plain colour is left out, a transparent one is none, a partly transparent one is opaque; each loss is named', () => {
  const line = { color: '000000', width: 9525 };
  // Patterns, textures, pictures and the kinds of gradients.
  for (const fillType of [1, 2, 3, 4, 5, 6, 7, 8, 9, 100]) assert.deepEqual(officeArtLook({ type: 1, fillType, fillColor: 0x0000ff }), { geometry: 'rect', line, lost: ['fill'] });
  assert.deepEqual(officeArtLook({ type: 1, fillType: 0, fillColor: 0x0000ff }), { geometry: 'rect', fill: 'FF0000', line });
  // Opacity is a 16.16 fixed point number.
  assert.deepEqual(officeArtLook({ type: 1, fillColor: 0x0000ff, fillOpacity: 0 }), { geometry: 'rect', line });
  assert.deepEqual(officeArtLook({ type: 1, fillColor: 0x0000ff, fillOpacity: -5 }), { geometry: 'rect', line });
  assert.deepEqual(officeArtLook({ type: 1, fillType: 7, fillOpacity: 0 }), { geometry: 'rect', line });
  assert.deepEqual(officeArtLook({ type: 1, fillColor: 0x0000ff, fillOpacity: 32768 }), { geometry: 'rect', fill: 'FF0000', line, lost: ['opacity'] });
  assert.deepEqual(officeArtLook({ type: 1, fillColor: 0x0000ff, fillOpacity: 65536 }), { geometry: 'rect', fill: 'FF0000', line });
  assert.deepEqual(officeArtLook({ type: 1, fillColor: 0x0000ff, fillOpacity: 70000 }), { geometry: 'rect', fill: 'FF0000', line });
  // Nothing of a fill that is switched off, or of a line, is lost.
  assert.deepEqual(officeArtLook({ type: 1, fFilled: false, fillType: 7, fillOpacity: 100 }), { geometry: 'rect', line });
  assert.deepEqual(officeArtLook({ type: 20, fillType: 7, fillOpacity: 100, fillColor: 0x08000001 }), { geometry: 'line', line });
});

// ------------------------------------------------------------------- reading

test('shapes of a doc file: kind, fill, outline, dashes, rounded corners, flipping', () => {
  const model = shapesDoc([
    { type: 1, rect: [0, 0, 1500, 600] },
    { type: 3, rect: [0, 0, 1500, 600], props: [[P_FILL_COLOR, 0x00ffff], [P_FILL_FLAGS, FILLED], [P_LINE_COLOR, 0xff0000], [P_LINE_WIDTH, 19050], [P_LINE_DASHING, 2], [P_LINE_FLAGS, LINED]] },
    { type: 2, rect: [0, 0, 1500, 600], props: [[P_ADJUST, 5400], [P_FILL_FLAGS, NOT_FILLED], [P_LINE_FLAGS, NOT_LINED]] },
    { type: 20, flags: 0x80, rect: [0, 0, 1500, 600], props: [[P_LINE_WIDTH, 0]] },
    { type: 32, flags: 0x40, rect: [0, 0, 1500, 0], props: [[P_ROTATION, 90 << 16]] },
    { type: 202, rect: [0, 0, 1500, 600], text: 'Boxed', props: [[P_FILL_FLAGS, NOT_FILLED], [P_LINE_WIDTH, 12700]] },
    { type: 13, rect: [0, 0, 1500, 600], props: [[P_FILL_COLOR, 0x112233]] },
  ]).read();
  assert.deepEqual(looks(model), [
    { geometry: 'rect', fill: 'FFFFFF', line: { color: '000000', width: 9525 } },
    { geometry: 'ellipse', fill: 'FFFF00', line: { color: '0000FF', width: 19050, dash: 'sysDot' } },
    { geometry: 'roundRect', corner: 25000 },
    { geometry: 'line', line: { color: '000000', width: 0 } },
    { geometry: 'line', line: { color: '000000', width: 9525 } },
    { geometry: 'rect', line: { color: '000000', width: 12700 } },
    { geometry: 'other', fill: '332211', line: { color: '000000', width: 9525 } },
  ]);
  const drawings = drawingsIn(model);
  assert.deepEqual(drawings[3].transform, { flipV: true });
  assert.deepEqual(drawings[4].transform, { rotation: 90 * 60000, flipH: true });
  assert.deepEqual([drawings[4].width, drawings[4].height], [1500 * 635, 0]);
  assert.equal(drawings[5].textBoxes!.length, 1);
  assert.ok(drawings.every((item) => item.kind === 'shape'));
});

test('shapes of a doc file: a flag counts only where its group says it is stated; a damaged table leaves the defaults', () => {
  const model = shapesDoc([
    // The values say "off" but the bits that state them are not set.
    { type: 1, rect: [0, 0, 100, 100], props: [[P_FILL_FLAGS, 0x000000], [P_LINE_FLAGS, 0x000000]] },
    // Other flags of the groups are stated, these are not.
    { type: 1, rect: [0, 0, 100, 100], props: [[P_FILL_FLAGS, 0x010001], [P_LINE_FLAGS, 0x100010]] },
    // A fill that is a gradient, half transparent: the gradient is not shown, and that is what is lost.
    { type: 1, rect: [0, 0, 100, 100], props: [[P_FILL_TYPE, 7], [P_FILL_COLOR, 0x0000ff], [P_FILL_OPACITY, 32768]] },
    // A plain fill, half transparent, in a colour of a scheme; an outline in a colour of the system.
    { type: 1, rect: [0, 0, 100, 100], props: [[P_FILL_COLOR, 0x0000ff], [P_FILL_OPACITY, 32768]] },
    { type: 1, rect: [0, 0, 100, 100], props: [[P_FILL_COLOR, 0x08000009], [P_LINE_COLOR, 0x100000f4]] },
    // A table that counts more properties than it has room for.
    { type: 3, rect: [0, 0, 100, 100], table: cat(u16le(3 | (4000 << 4)), u16le(0xf00b), u32le(6), u16le(P_LINE_WIDTH), u32le(25400)) },
    // A table cut off in the middle of a property.
    { type: 3, rect: [0, 0, 100, 100], table: cat(u16le(3 | (2 << 4)), u16le(0xf00b), u32le(9), u16le(P_LINE_WIDTH), u32le(25400), u8s(1, 2, 3)) },
  ]).read();
  const plain = { geometry: 'rect', fill: 'FFFFFF', line: { color: '000000', width: 9525 } };
  assert.deepEqual(looks(model), [
    plain,
    plain,
    { geometry: 'rect', line: plain.line, lost: ['fill'] },
    { ...plain, fill: 'FF0000', lost: ['opacity'] },
    { geometry: 'rect', line: plain.line, lost: ['color'] },
    { geometry: 'ellipse', fill: 'FFFFFF', line: { color: '000000', width: 25400 } },
    { geometry: 'ellipse', fill: 'FFFFFF', line: { color: '000000', width: 25400 } },
  ]);
});

test('shapes of a doc file: a shape the drawing does not have gets no look', () => {
  const builder = shapesDoc([{ type: 1, rect: [0, 0, 100, 100] }]);
  const model = builder.read({
    patch: ({ table, fcLcb }) => {
      // The character names a shape the drawing does not have.
      const [at] = fcLcb(FC.PlcSpaMom);
      new DataView(table.buffer, table.byteOffset).setUint32(at + 8, 4242, true);
    },
  });
  const [drawing] = drawingsIn(model);
  assert.equal(drawing.kind, 'shape');
  assert.equal(drawing.shape, undefined);
});

test('shapes of an RTF file: kind, fill, outline, dashes, rounded corners, flipping', async () => {
  const model = await rtfModel(
    'Anchor ' +
      shp([0, 0, 1500, 600], sp('shapeType', 1)) +
      shp([0, 0, 1500, 600], sp('shapeType', 3) + sp('fillColor', 65535) + sp('fFilled', 1) + sp('lineColor', 16711680) + sp('lineWidth', 19050) + sp('lineDashing', 2) + sp('fLine', 1)) +
      shp([0, 0, 1500, 600], sp('shapeType', 2) + sp('adjustValue', 5400) + sp('fFilled', 0) + sp('fLine', 0)) +
      shp([0, 0, 1500, 600], sp('shapeType', 20) + sp('fFlipV', 1) + sp('lineWidth', 0)) +
      shp([0, 0, 1500, 0], sp('shapeType', 32) + sp('fFlipH', 1) + sp('rotation', 90 * 65536)) +
      shp([0, 0, 1500, 600], sp('shapeType', 202) + sp('fFilled', 0) + sp('lineWidth', 12700), 'Boxed') +
      shp([0, 0, 1500, 600], sp('shapeType', 13) + sp('fillColor', 0x112233)),
  );
  assert.deepEqual(looks(model), [
    { geometry: 'rect', fill: 'FFFFFF', line: { color: '000000', width: 9525 } },
    { geometry: 'ellipse', fill: 'FFFF00', line: { color: '0000FF', width: 19050, dash: 'sysDot' } },
    { geometry: 'roundRect', corner: 25000 },
    { geometry: 'line', line: { color: '000000', width: 0 } },
    { geometry: 'line', line: { color: '000000', width: 9525 } },
    { geometry: 'rect', line: { color: '000000', width: 12700 } },
    { geometry: 'other', fill: '332211', line: { color: '000000', width: 9525 } },
  ]);
  const drawings = drawingsIn(model);
  assert.deepEqual(drawings[3].transform, { flipV: true });
  assert.deepEqual(drawings[4].transform, { rotation: 90 * 60000, flipH: true });
  assert.equal(drawings[5].textBoxes!.length, 1);
});

test('shapes of an RTF file: values that are no numbers leave the defaults; pictures, groups and Word 95 objects get no look', async () => {
  const model = await rtfModel(
    'Anchor ' +
      shp([0, 0, 100, 100], sp('shapeType', 'oval') + sp('fillColor', 'red') + sp('lineWidth', 'thick') + sp('lineDashing', '-') + sp('fFilled', '') + sp('fLine', 'x')) +
      shp([0, 0, 100, 100], sp('lineWidth', '-12700') + sp('lineDashing', 77)) +
      // A gradient; a half transparent fill; colours that are indexes.
      shp([0, 0, 100, 100], sp('fillType', 7) + sp('fillColor', 255)) +
      shp([0, 0, 100, 100], sp('fillColor', 255) + sp('fillOpacity', 32768)) +
      shp([0, 0, 100, 100], sp('fillColor', 0x08000009) + sp('lineColor', '99999999999999999999')) +
      shp([0, 0, 100, 100], sp('fillColor', 255) + sp('fillOpacity', 0)) +
      // A group: what its parts look like is not in the model.
      `{\\shpgrp{\\*\\shpinst\\shpleft0\\shptop0\\shpright100\\shpbottom100${shp([0, 0, 50, 50], sp('shapeType', 3))}}}` +
      // A drawing object of Word 95 with text.
      '{\\*\\do\\dobxcolumn\\dobypara\\dptxbx{\\dptxbxtext Old\\par}\\dpx0\\dpy0\\dpxsize500\\dpysize500\\dplinew15\\dpfillpat1}',
  );
  const drawings = drawingsIn(model);
  const plain = { geometry: 'rect', fill: 'FFFFFF', line: { color: '000000', width: 9525 } };
  assert.deepEqual(drawings.map((item) => [item.kind, item.shape]), [
    ['shape', plain],
    ['shape', plain],
    ['shape', { geometry: 'rect', line: plain.line, lost: ['fill'] }],
    ['shape', { ...plain, fill: 'FF0000', lost: ['opacity'] }],
    ['shape', { geometry: 'rect', line: plain.line, lost: ['color'] }],
    ['shape', { geometry: 'rect', line: plain.line }],
    ['group', undefined],
    ['shape', undefined],
  ]);
});

// -------------------------------------------------------------------- layout

const RTF_SHAPES =
  'Text ' +
  // A red line, 2 px thick, 200 px long, 20 px below the paragraph.
  shp([0, 300, 3000, 300], sp('shapeType', 20) + sp('lineColor', 255) + sp('lineWidth', 19050)) +
  // A text box filled in yellow with a blue outline.
  shp([0, 600, 1500, 1200], sp('shapeType', 202) + sp('fillColor', 0x00ffff) + sp('lineColor', 0xff0000) + sp('lineWidth', 12700), 'Boxed') +
  shp([1800, 600, 3000, 1200], sp('shapeType', 3) + sp('fFilled', 0) + sp('lineDashing', 6)) +
  shp([0, 1500, 1500, 2100], sp('shapeType', 2) + sp('fLine', 0) + sp('fillColor', 0x00ff00)) +
  shp([1800, 1500, 3000, 2100], sp('shapeType', 20) + sp('fFlipV', 1)) +
  shp([0, 2400, 1500, 3000], sp('shapeType', 1) + sp('rotation', 45 * 65536) + sp('lineWidth', 0)) +
  shp([1800, 2400, 3000, 3000], sp('shapeType', 3) + sp('fillColor', 0x0000ff) + sp('lineDashing', 1), 'In');

test('shapes on the page: lines, boxes, ellipses and rounded rectangles with their fill and outline', async () => {
  const page = pageOf(await rtfModel(RTF_SHAPES));
  assert.equal(placeholders(page).length, 0);
  assert.deepEqual(shapesOf(page), [
    { type: 'shape', x: 40, y: 60, width: 200, height: 0, geometry: 'line', stroke: { color: '#FF0000', width: 2 } },
    { type: 'shape', x: 160, y: 80, width: 80, height: 40, geometry: 'ellipse', stroke: { color: '#000000', width: 1, dash: 'dash' } },
    // Corners of a sixth of the shorter side.
    { type: 'shape', x: 40, y: 140, width: 100, height: 40, geometry: 'roundRect', radius: 40 * 0.16667, fill: '#00FF00' },
    { type: 'shape', x: 160, y: 140, width: 80, height: 40, geometry: 'line', flipV: true, stroke: { color: '#000000', width: 1 } },
    // A box that is turned is a shape the painter turns; an outline of no width is the thinnest line.
    { type: 'shape', x: 40, y: 200, width: 100, height: 40, geometry: 'rect', rotation: 45, fill: '#FFFFFF', stroke: { color: '#000000', width: 0.5 } },
    // A shape with text: its fill under the text, its outline over it.
    { type: 'shape', x: 160, y: 200, width: 80, height: 40, geometry: 'ellipse', fill: '#FF0000' },
    { type: 'shape', x: 160, y: 200, width: 80, height: 40, geometry: 'ellipse', stroke: { color: '#000000', width: 1, dash: 'sysDash' } },
  ]);
  // The text box is a filled rectangle and four lines, as that of a docx is.
  assert.deepEqual(rects(page), [{ type: 'rect', x: 40, y: 80, width: 100, height: 40, fill: '#FFFF00' }]);
  assert.deepEqual(
    lines(page).map((line) => [line.x1, line.y1, line.x2, line.y2, line.color, Math.round(line.width * 100) / 100, line.style]),
    [
      [40, 80, 140, 80, '#0000FF', 1.33, 'single'],
      [40, 120, 140, 120, '#0000FF', 1.33, 'single'],
      [40, 80, 40, 120, '#0000FF', 1.33, 'single'],
      [140, 80, 140, 120, '#0000FF', 1.33, 'single'],
    ],
  );
  const order = (item: PageItem | undefined) => page.items.indexOf(item!);
  const boxed = texts(page).find((item) => item.text === 'Boxed');
  assert.ok(order(rects(page)[0]) < order(boxed) && order(boxed) < order(lines(page)[0]));
  const inside = texts(page).find((item) => item.text === 'In');
  const ellipses = shapesOf(page).filter((item) => item.x === 160 && item.y === 200);
  assert.ok(order(ellipses[0]) < order(inside) && order(inside) < order(ellipses[1]));
});

test('shapes on the page: a shape of a doc file is drawn as that of an RTF file', async () => {
  const doc = shapesDoc([
    { type: 20, rect: [0, 300, 3000, 300], props: [[P_LINE_COLOR, 255], [P_LINE_WIDTH, 19050]] },
    { type: 202, rect: [0, 600, 1500, 1200], text: 'Boxed', props: [[P_FILL_COLOR, 0x00ffff], [P_LINE_COLOR, 0xff0000], [P_LINE_WIDTH, 12700]] },
    { type: 3, rect: [1800, 600, 3000, 1200], props: [[P_FILL_FLAGS, NOT_FILLED], [P_LINE_DASHING, 6]] },
  ]).read();
  const page = pageOf(doc);
  const [line, ellipse] = shapesOf(page);
  assert.deepEqual([line.geometry, line.width, line.height, line.stroke], ['line', 200, 0, { color: '#FF0000', width: 2 }]);
  assert.deepEqual([ellipse.geometry, ellipse.width, ellipse.height, ellipse.fill, ellipse.stroke], ['ellipse', 80, 40, undefined, { color: '#000000', width: 1, dash: 'dash' }]);
  assert.equal(ellipse.x - line.x, 120);
  assert.equal(ellipse.y - line.y, 20);
  assert.deepEqual(rects(page).map((rect) => [rect.width, rect.height, rect.fill]), [[100, 40, '#FFFF00']]);
  assert.deepEqual([lines(page).length, lines(page)[0].color], [4, '#0000FF']);
});

test('shapes on the page: nothing to show paints nothing; a geometry the reader does not know is a placeholder, or a box around its text', async () => {
  const page = pageOf(
    await rtfModel(
      'Text ' +
        // Neither filled nor outlined.
        shp([0, 300, 1500, 900], sp('shapeType', 3) + sp('fFilled', 0) + sp('fLine', 0)) +
        // A line of no length.
        shp([0, 300, 0, 300], sp('shapeType', 20)) +
        // An arrow without text, and one with.
        shp([0, 1200, 1500, 1800], sp('shapeType', 13) + sp('fillColor', 255)) +
        shp([1800, 1200, 3300, 1800], sp('shapeType', 13) + sp('fillColor', 255) + sp('fLine', 0), 'Arrow'),
    ),
  );
  assert.deepEqual(shapesOf(page), []);
  assert.deepEqual(placeholders(page).map((item) => [item.placeholder, item.x, item.y, item.width, item.height]), [['shape', 40, 120, 100, 40]]);
  assert.deepEqual(rects(page), [{ type: 'rect', x: 160, y: 120, width: 100, height: 40, fill: '#FF0000' }]);
  assert.equal(lines(page).length, 0);
});

test('shapes on the page: a look that is not what the model promises is not trusted', () => {
  const item = (shape: unknown): DrawingItem => ({
    type: 'drawing',
    placement: 'inline',
    kind: 'shape',
    width: 952500,
    height: 476250,
    shape: shape as ShapeLook,
    xml: { type: 'element', name: 'doc:drawing', attrs: new Map(), children: [] },
  });
  const model = shapesDoc([]).read();
  const paragraph = model.body.blocks[0] as Paragraph;
  paragraph.children.push({
    type: 'run',
    props: {},
    content: [
      item({ geometry: 'ellipse', fill: 'red"; }', line: { color: 'ZZZZZZ', width: 9525 } }),
      item({ geometry: 'ellipse', fill: 'abcdef', line: { color: '00ff00', width: Infinity, dash: 'constructor' } }),
      item({ geometry: 'roundRect', corner: NaN, line: { color: '000000', width: -1 } }),
    ],
  });
  const shapes = shapesOf(pageOf(model));
  assert.deepEqual(
    shapes.map((shape) => [shape.geometry, shape.fill, shape.stroke, shape.radius]),
    [
      ['ellipse', '#ABCDEF', { color: '#00FF00', width: 0.5 }, undefined],
      ['roundRect', undefined, { color: '#000000', width: 0.5 }, 0],
    ],
  );
});

test('shapes of a docx: an ellipse, a line and a rounded rectangle are drawn; what cannot be shown stays a placeholder', async () => {
  const drawing = (width: number, height: number, inner: string, transform = '') =>
    `<w:r><w:drawing><wp:inline><wp:extent cx="${width * 9525}" cy="${height * 9525}"/><wp:docPr id="1" name="s"/>` +
    '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/>' +
    `<wps:spPr>${transform}${inner}</wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  const geom = (prst: string, guides = '') => `<a:prstGeom prst="${prst}"><a:avLst>${guides}</a:avLst></a:prstGeom>`;
  const fill = (rgb: string) => `<a:solidFill><a:srgbClr val="${rgb}"/></a:solidFill>`;
  const { pages } = await lay(
    p(drawing(100, 50, geom('ellipse') + fill('FFEEDD') + `<a:ln w="19050">${fill('112233')}<a:prstDash val="lgDash"/></a:ln>`)) +
      p(drawing(100, 0, geom('line') + `<a:ln w="0">${fill('FF0000')}</a:ln>`, '<a:xfrm flipH="1" rot="5400000"><a:off x="0" y="0"/><a:ext cx="952500" cy="0"/></a:xfrm>')) +
      p(drawing(100, 50, geom('roundRect', '<a:gd name="adj" fmla="val 50000"/>') + fill('00FF00') + `<a:ln>${fill('0000FF')}<a:prstDash val="solid"/></a:ln>`)) +
      p(drawing(100, 50, geom('roundRect', '<a:gd name="adj" fmla="val 999999999"/>') + fill('00FF00'))) +
      // A dashed box is a shape; a solid one is a rectangle and four lines.
      p(drawing(100, 50, geom('rect') + `<a:ln>${fill('0000FF')}<a:prstDash val="sysDot"/></a:ln>`)) +
      p(drawing(100, 50, geom('rect') + fill('ABCDEF') + `<a:ln>${fill('0000FF')}</a:ln>`)) +
      // Not drawn: a gradient, a geometry that is not known, a custom geometry, a shape that states no geometry.
      p(drawing(100, 50, geom('rect') + '<a:gradFill/>')) +
      p(drawing(100, 50, geom('star5') + fill('ABCDEF'))) +
      p(drawing(100, 50, '<a:custGeom/>' + fill('ABCDEF'))) +
      p(drawing(100, 50, fill('ABCDEF'))),
  );
  // The drawings run over several pages.
  const page: LayoutPage = { ...pages[0], items: pages.flatMap((each) => each.items) };
  assert.deepEqual(
    shapesOf(page).map((shape) => [shape.geometry, shape.width, shape.height, shape.fill, shape.stroke, shape.radius, shape.rotation, shape.flipH]),
    [
      ['ellipse', 100, 50, '#FFEEDD', { color: '#112233', width: 2, dash: 'lgDash' }, undefined, undefined, undefined],
      ['line', 100, 0, undefined, { color: '#FF0000', width: 0.5 }, undefined, 90, true],
      ['roundRect', 100, 50, '#00FF00', { color: '#0000FF', width: 1 }, 25, undefined, undefined],
      ['roundRect', 100, 50, '#00FF00', undefined, 50 * 0.16667, undefined, undefined],
      ['rect', 100, 50, undefined, { color: '#0000FF', width: 1, dash: 'sysDot' }, undefined, undefined, undefined],
    ],
  );
  assert.deepEqual(rects(page).map((rect) => [rect.width, rect.height, rect.fill]), [[100, 50, '#ABCDEF']]);
  assert.equal(lines(page).length, 4);
  assert.deepEqual(placeholders(page).map((item) => item.placeholder), ['shape', 'shape', 'shape', 'shape']);
});

test('a shape that is to show a picture the file does not have stays a picture: a placeholder on the page, counted when converted', async () => {
  const doc = shapesDoc([
    // A picture frame without a picture, and a box that names a picture the store does not have.
    { type: 75, rect: [0, 0, 1500, 600] },
    { type: 1, rect: [0, 900, 1500, 1500], props: [[P_PICTURE, 3]] },
    // The same with text: a text box, without a look that would fill it in white.
    { type: 75, rect: [0, 1800, 1500, 2400], text: 'Framed' },
  ]).build();
  const source = rtf(
    'Text ' +
      shp([0, 0, 1500, 600], sp('shapeType', 75)) +
      shp([0, 900, 1500, 1500], sp('shapeType', 1) + '{\\sp{\\sn pib}{\\sv {\\pict\\pngblip\\picwgoal600\\pichgoal300 }}}') +
      shp([0, 1800, 1500, 2400], sp('shapeType', 75), 'Framed'),
  );
  for (const bytes of [doc, source]) {
    const file = await open(bytes);
    const model = await file.document!();
    assert.deepEqual(drawingsIn(model).map((item) => [item.kind, item.shape, item.image, item.textBoxes ? item.textBoxes.length : 0]), [
      ['picture', undefined, undefined, 0],
      ['picture', undefined, undefined, 0],
      ['shape', undefined, undefined, 1],
    ]);
    const page = pageOf(model);
    assert.deepEqual(placeholders(page).map((item) => [item.placeholder, item.width, item.height]), [
      ['picture', 100, 40],
      ['picture', 100, 40],
    ]);
    assert.deepEqual([shapesOf(page), rects(page), lines(page)], [[], [], []]);
    const converted = await convertForEditing(file);
    assert.deepEqual(converted!.warnings, [
      '1 shape(s) were written as plain rectangles: their outline, fill and geometry were not carried over',
      '2 picture(s) whose data is not available were written as empty frames',
    ]);
  }
});

test('shapes on the page and converted: a fill that is not a plain colour is not shown, a colour that is an index is not made up; each is counted', async () => {
  const file = await open(
    rtf(
      'Text ' +
        // A gradient without text: a placeholder, as in a docx. With text: the outline, no fill.
        shp([0, 300, 1500, 900], sp('shapeType', 3) + sp('fillType', 7) + sp('fillColor', 255)) +
        shp([1800, 300, 3300, 900], sp('shapeType', 1) + sp('fillType', 1) + sp('fillColor', 255), 'Pattern') +
        // Half transparent: shown, and written, opaque.
        shp([0, 1200, 1500, 1800], sp('shapeType', 1) + sp('fillColor', 65280) + sp('fillOpacity', 32768) + sp('fLine', 0), 'Glass') +
        // Fully transparent: no fill, nothing lost.
        shp([1800, 1200, 3300, 1800], sp('shapeType', 3) + sp('fillColor', 255) + sp('fillOpacity', 0)) +
        // A fill in a colour of a scheme: not painted; an outline in a system colour: black.
        shp([0, 2100, 1500, 2700], sp('shapeType', 3) + sp('fillColor', 0x08000009) + sp('lineColor', 0x100000f4)) +
        shp([1800, 2100, 3300, 2700], sp('shapeType', 3) + sp('lineColor', 0x01000002) + sp('fFilled', 0)),
    ),
  );
  const page = pageOf(await file.document!());
  assert.deepEqual(placeholders(page).map((item) => [item.placeholder, item.x, item.y]), [['shape', 40, 60]]);
  assert.deepEqual(rects(page).map((rect) => [rect.x, rect.y, rect.fill]), [[40, 120, '#00FF00']]);
  // The outline of the box with the pattern.
  assert.deepEqual([lines(page).length, lines(page)[0].x1, lines(page)[0].y1], [4, 160, 60]);
  assert.deepEqual(
    shapesOf(page).map((item) => [item.x, item.y, item.geometry, item.fill, item.stroke && item.stroke.color]),
    [
      [160, 120, 'ellipse', undefined, '#000000'],
      [40, 180, 'ellipse', undefined, '#000000'],
      [160, 180, 'ellipse', undefined, '#000000'],
    ],
  );
  const converted = await convertForEditing(file);
  assert.deepEqual(converted!.warnings, [
    '2 shape(s) with a fill that is not a plain colour (a pattern, a picture, a gradient) were written without a fill',
    '1 shape(s) with a partly transparent fill were written with an opaque one',
    '2 shape(s) with a colour that is an index into a palette or a system colour were written without that colour: no fill, a black outline',
  ]);
});

test('shapes of a docx: the first shape of a group says nothing about the group', async () => {
  const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
  const WPG = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';
  const xfrm = (x: number, y: number, w: number, h: number) => `<a:xfrm><a:off x="${x * 9525}" y="${y * 9525}"/><a:ext cx="${w * 9525}" cy="${h * 9525}"/></a:xfrm>`;
  const part = (x: number, y: number, w: number, h: number, look: string, text = '') =>
    `<wps:wsp><wps:cNvSpPr/><wps:spPr>${xfrm(x, y, w, h)}${look}</wps:spPr>${text ? `<wps:txbx><w:txbxContent>${p(text)}</w:txbxContent></wps:txbx>` : ''}<wps:bodyPr/></wps:wsp>`;
  const red = '<a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>';
  const green = '<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>';
  const geom = (prst: string) => `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>`;
  const group = (parts: string) =>
    '<w:r><w:drawing><wp:inline><wp:extent cx="1905000" cy="952500"/><wp:docPr id="1" name="g"/>' +
    `<a:graphic><a:graphicData uri="${WPG}"><wpg:wgp xmlns:wpg="${WPG}" xmlns:wps="${WPS}"><wpg:cNvGrpSpPr/>` +
    `<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/><a:chOff x="0" y="0"/><a:chExt cx="1905000" cy="952500"/></a:xfrm></wpg:grpSpPr>` +
    `${parts}</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  // A flow chart: a connector first, then two boxes with text.
  const chart = group(part(60, 20, 80, 0, geom('straightConnector1') + red) + part(0, 0, 60, 40, geom('rect') + green + red, 'Start') + part(140, 0, 60, 40, geom('ellipse') + green, 'End'));
  for (const first of ['straightConnector1', 'line', 'ellipse', 'roundRect', 'rect']) {
    const body = p(chart.replace('straightConnector1', first)) + p(group(part(0, 0, 40, 40, geom(first) + green + red)));
    const { pages, doc } = await lay(body);
    assert.deepEqual(drawingsIn(doc).map((item) => item.kind), ['group', 'group'], first);
    const page = pages[0];
    // Nothing of the first part is drawn over the 200 x 100 px of the group; its text is there.
    assert.deepEqual([shapesOf(page), rects(page), lines(page)], [[], [], []], first);
    assert.deepEqual(texts(page).map((item) => item.text), ['Start', 'End'], first);
    // The group without text is a placeholder of its size.
    assert.deepEqual(placeholders(page).map((item) => [item.placeholder, item.width, item.height]), [['group', 200, 100]], first);
  }
});

test('shapes of a docx: colours of the shape style, an outline that is switched off', async () => {
  const drawing = (inner: string) =>
    '<w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><wp:docPr id="1" name="s"/>' +
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/>${inner}<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  const ellipse = '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>';
  const style = (line: string, fill: string) =>
    `<wps:style><a:lnRef idx="${line}"><a:srgbClr val="00AA00"/></a:lnRef><a:fillRef idx="${fill}"><a:srgbClr val="EEEEEE"/></a:fillRef>` +
    '<a:effectRef idx="0"><a:srgbClr val="000000"/></a:effectRef><a:fontRef idx="minor"><a:srgbClr val="000000"/></a:fontRef></wps:style>';
  const solid = (rgb: string) => `<a:solidFill><a:srgbClr val="${rgb}"/></a:solidFill>`;
  const { pages } = await lay(
    // The style gives what the properties leave open.
    p(drawing(`<wps:spPr>${ellipse}</wps:spPr>` + style('2', '1'))) +
      // An index of 0 is "none".
      p(drawing(`<wps:spPr>${ellipse}<a:ln w="9525"/></wps:spPr>` + style('0', '1'))) +
      p(drawing(`<wps:spPr>${ellipse}</wps:spPr>` + style('1', '0'))) +
      // What the properties state wins: no fill, no outline; an outline of the width 0 is the thinnest line.
      p(drawing(`<wps:spPr>${ellipse}<a:noFill/><a:ln w="0">${solid('FF0000')}</a:ln></wps:spPr>` + style('2', '1'))) +
      p(drawing(`<wps:spPr>${ellipse}${solid('0000FF')}<a:ln><a:noFill/></a:ln></wps:spPr>` + style('2', '1'))) +
      // Nothing to show at all: nothing is drawn, and no placeholder stands there.
      p(drawing(`<wps:spPr>${ellipse}<a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>` + style('2', '1'))) +
      // An outline that is a gradient is not drawn.
      p(drawing(`<wps:spPr>${ellipse}${solid('0000FF')}<a:ln w="19050"><a:gradFill/></a:ln></wps:spPr>`)),
  );
  const page: LayoutPage = { ...pages[0], items: pages.flatMap((each) => each.items) };
  assert.deepEqual(
    shapesOf(page).map((item) => [item.fill, item.stroke]),
    [
      ['#EEEEEE', { color: '#00AA00', width: 1 }],
      ['#EEEEEE', undefined],
      [undefined, { color: '#00AA00', width: 1 }],
      [undefined, { color: '#FF0000', width: 0.5 }],
      ['#0000FF', undefined],
      ['#0000FF', undefined],
    ],
  );
  assert.equal(placeholders(page).length, 0);
});

// ------------------------------------------------------------------ painting

interface PathCall {
  path: string[];
  fill?: string;
  stroke?: string;
  width?: number;
  dash?: number[];
  /** The transform the path was made under. */
  matrix: number[];
}

/** A canvas that records the paths that are filled and stroked, and the rectangles that are filled. */
function recorder(dash = true) {
  const paths: PathCall[] = [];
  const fills: { rect: number[]; fill: string }[] = [];
  let m = [1, 0, 0, 1, 0, 0];
  const stack: number[][] = [];
  let path: string[] = [];
  let lineDash: number[] = [];
  const multiply = (n: number[]) => {
    m = [m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1], m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3], m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5]];
  };
  const round = (value: number) => Math.round(value * 100) / 100;
  const point = (x: number, y: number) => `${round(x)},${round(y)}`;
  const ctx: { [name: string]: unknown } = {
    fillStyle: '#000000',
    strokeStyle: '#000000',
    lineWidth: 1,
    save: () => stack.push(m.slice()),
    restore: () => {
      m = stack.pop() || m;
    },
    scale: (x: number, y: number) => multiply([x, 0, 0, y, 0, 0]),
    translate: (x: number, y: number) => multiply([1, 0, 0, 1, x, y]),
    rotate: (angle: number) => multiply([Math.cos(angle), Math.sin(angle), -Math.sin(angle), Math.cos(angle), 0, 0]),
    beginPath: () => {
      path = [];
    },
    moveTo: (x: number, y: number) => path.push('M' + point(x, y)),
    lineTo: (x: number, y: number) => path.push('L' + point(x, y)),
    bezierCurveTo: (a: number, b: number, c: number, d: number, x: number, y: number) => path.push('C' + point(x, y)),
    closePath: () => path.push('Z'),
    fill: () => paths.push({ path: path.slice(), fill: ctx.fillStyle as string, matrix: m.map(round) }),
    stroke: () => paths.push({ path: path.slice(), stroke: ctx.strokeStyle as string, width: ctx.lineWidth as number, dash: lineDash.slice(), matrix: m.map(round) }),
    fillRect: (x: number, y: number, width: number, height: number) => fills.push({ rect: [x, y, width, height].map(round), fill: ctx.fillStyle as string }),
  };
  if (dash)
    ctx.setLineDash = (value: number[]) => {
      lineDash = value.slice();
    };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, paths, fills, depth: () => stack.length };
}

function paint(items: PageItem[], scale = 1, dash = true) {
  const rec = recorder(dash);
  renderPage(rec.ctx, { width: 400, height: 300, section: 0, number: 1, items }, { background: null, scale });
  assert.equal(rec.depth(), 0, 'every save has its restore');
  return rec;
}

const shape = (extra: Partial<ShapeItem>): ShapeItem => ({ type: 'shape', x: 20, y: 30, width: 100, height: 40, geometry: 'rect', ...extra });

test('painting shapes: an ellipse and a rounded rectangle are filled and then outlined around their centre', () => {
  const rec = paint([shape({ geometry: 'ellipse', fill: '#ffeedd', stroke: { color: '#112233', width: 2 } })]);
  const outline = ['M50,0', 'C0,20', 'C-50,0', 'C0,-20', 'C50,0', 'Z'];
  assert.deepEqual(rec.paths, [
    { path: outline, fill: '#ffeedd', matrix: [1, 0, 0, 1, 70, 50] },
    { path: outline, stroke: '#112233', width: 2, dash: [], matrix: [1, 0, 0, 1, 70, 50] },
  ]);
  const rounded = paint([shape({ geometry: 'roundRect', radius: 10, fill: '#00ff00' })]);
  assert.deepEqual(rounded.paths, [
    { path: ['M-40,-20', 'L40,-20', 'C50,-10', 'L50,10', 'C40,20', 'L-40,20', 'C-50,10', 'L-50,-10', 'C-40,-20', 'Z'], fill: '#00ff00', matrix: [1, 0, 0, 1, 70, 50] },
  ]);
  // Corners larger than half a side are cut to it; without a radius the shape is a box.
  assert.equal(paint([shape({ geometry: 'roundRect', radius: 500, fill: '#00ff00' })]).paths[0].path[0], 'M-30,-20');
  assert.deepEqual(paint([shape({ geometry: 'roundRect', fill: '#00ff00' })]).paths[0].path, ['M-50,-20', 'L50,-20', 'L50,20', 'L-50,20', 'Z']);
});

test('painting shapes: dashes scale with the outline; turning and flipping', () => {
  const dashed = paint([shape({ stroke: { color: '#000000', width: 2, dash: 'dash' } })]);
  assert.deepEqual([dashed.paths[0].width, dashed.paths[0].dash], [2, [8, 6]]);
  // A dash that is not known, or a canvas that cannot dash: a solid line.
  assert.deepEqual(paint([shape({ stroke: { color: '#000000', width: 2, dash: 'toString' } })]).paths[0].dash, []);
  assert.equal(paint([shape({ stroke: { color: '#000000', width: 2, dash: 'dash' } })], 1, false).paths.length, 1);
  // The thinnest line is one pixel of the canvas, whatever the zoom.
  assert.equal(paint([shape({ stroke: { color: '#000000', width: 0.5 } })], 0.25).paths[0].width, 4);
  // Turned by 90 degrees around its centre.
  assert.deepEqual(paint([shape({ fill: '#ff0000', rotation: 90 })]).paths[0].matrix, [0, 1, -1, 0, 70, 50]);
  // A line runs from the top left to the bottom right; flipped once it runs upwards, flipped twice it is the same line.
  const line = (extra: Partial<ShapeItem>) => paint([shape({ geometry: 'line', fill: '#ff0000', stroke: { color: '#0000ff', width: 1 }, ...extra })]).paths;
  assert.deepEqual(line({}), [{ path: ['M-50,-20', 'L50,20'], stroke: '#0000ff', width: 1, dash: [], matrix: [1, 0, 0, 1, 70, 50] }]);
  assert.deepEqual(line({ flipV: true })[0].path, ['M-50,20', 'L50,-20']);
  assert.deepEqual(line({ flipH: true })[0].path, ['M-50,20', 'L50,-20']);
  assert.deepEqual(line({ flipH: true, flipV: true })[0].path, ['M-50,-20', 'L50,20']);
});

test('painting shapes: a horizontal or vertical line lies on whole pixels; what is not a shape paints nothing', () => {
  const flat = paint([shape({ geometry: 'line', height: 0, y: 30.4, stroke: { color: '#ff0000', width: 2 } })]);
  assert.deepEqual(flat.paths, []);
  assert.deepEqual(flat.fills, [{ rect: [20, 29, 100, 2], fill: '#ff0000' }]);
  const upright = paint([shape({ geometry: 'line', width: 0, stroke: { color: '#ff0000', width: 1 } })]);
  assert.deepEqual(upright.fills, [{ rect: [20, 30, 1, 40], fill: '#ff0000' }]);
  // Turned, it is a path again.
  assert.equal(paint([shape({ geometry: 'line', height: 0, rotation: 30, stroke: { color: '#ff0000', width: 2 } })]).paths.length, 1);
  const nothing = paint([
    shape({}),
    shape({ fill: 'red' }),
    shape({ width: 0, fill: '#ff0000' }),
    shape({ width: NaN, fill: '#ff0000' }),
    shape({ x: Infinity, fill: '#ff0000' }),
    shape({ height: -4, fill: '#ff0000' }),
    shape({ geometry: 'line', width: 0, height: 0, stroke: { color: '#ff0000', width: 1 } }),
    shape({ geometry: 'line', fill: '#ff0000' }),
    shape({ stroke: { color: '#ff0000', width: NaN } }),
  ]);
  assert.deepEqual([nothing.paths, nothing.fills], [[], []]);
  // A colour that is none is black, a width beyond reason is cut.
  const odd = paint([shape({ stroke: { color: 'url(x)', width: 5000 } })]).paths[0];
  assert.deepEqual([odd.stroke, odd.width], ['#000000', 64]);
});

// ------------------------------------------------------------------- writing

async function documentXml(bytes: Uint8Array): Promise<string> {
  const pkg = await OpcPackage.open(bytes);
  return new TextDecoder().decode(await pkg.readPart('/word/document.xml'));
}

test('shapes written as docx: geometry, fill and outline in the order of the schema, without a warning', async () => {
  const converted = await convertForEditing(await open(rtf(RTF_SHAPES)));
  assert.deepEqual(converted!.warnings, []);
  const xml = await documentXml(converted!.bytes);
  const props = [...xml.matchAll(/<wps:wsp[^>]*>(<wps:cNv[^>]*\/>)<wps:spPr>(.*?)<\/wps:spPr>/g)].map((match) => match[1] + match[2].replace(/<a:xfrm.*?<\/a:xfrm>/, (xfrm) => /^<a:xfrm[^>]*>/.exec(xfrm)![0]));
  const ln = (width: number, rgb: string, dash = '') => `<a:ln w="${width}"><a:solidFill><a:srgbClr val="${rgb}"/></a:solidFill>${dash}</a:ln>`;
  const fill = (rgb: string) => `<a:solidFill><a:srgbClr val="${rgb}"/></a:solidFill>`;
  const geom = (prst: string) => `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>`;
  assert.deepEqual(props, [
    '<wps:cNvCnPr/><a:xfrm>' + geom('line') + '<a:noFill/>' + ln(19050, 'FF0000'),
    '<wps:cNvSpPr txBox="1"/><a:xfrm>' + geom('rect') + fill('FFFF00') + ln(12700, '0000FF'),
    '<wps:cNvSpPr/><a:xfrm>' + geom('ellipse') + '<a:noFill/>' + ln(9525, '000000', '<a:prstDash val="dash"/>'),
    '<wps:cNvSpPr/><a:xfrm>' + geom('roundRect') + fill('00FF00') + '<a:ln><a:noFill/></a:ln>',
    '<wps:cNvCnPr/><a:xfrm flipV="1">' + geom('line') + '<a:noFill/>' + ln(9525, '000000'),
    '<wps:cNvSpPr/><a:xfrm rot="2700000">' + geom('rect') + fill('FFFFFF') + ln(0, '000000'),
    '<wps:cNvSpPr txBox="1"/><a:xfrm>' + geom('ellipse') + fill('FF0000') + ln(9525, '000000', '<a:prstDash val="sysDash"/>'),
  ]);
});

test('shapes written as docx: the corners of a rounded rectangle; a look that is not what the model promises is written as nothing', async () => {
  const file = await open(rtf('Text ' + shp([0, 0, 1500, 600], sp('shapeType', 2) + sp('adjustValue', 5400)) + shp([0, 0, 1500, 600], sp('shapeType', 3)) + shp([0, 0, 1500, 600], sp('shapeType', 1))));
  const model = await file.document!();
  const [, ellipse, box] = drawingsIn(model);
  ellipse.shape = { geometry: 'ellipse', fill: '"/><x', line: { color: '00ff0', width: 5 } } as ShapeLook;
  box.shape = { geometry: 'rect', fill: 'abcdef', line: { color: 'abcdef', width: NaN, dash: '"/><x' } } as unknown as ShapeLook;
  const xml = await documentXml((await convertForEditing(file))!.bytes);
  assert.ok(xml.includes('<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom>'));
  assert.ok(xml.includes('<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>'));
  assert.ok(xml.includes('<a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:ln>'));
  assert.ok(!xml.includes('<x'));
});

test('shapes written as docx: only what is not carried over is counted in the warnings', async () => {
  const converted = await convertForEditing(
    await open(
      rtf(
        'Text ' +
          shp([0, 0, 1500, 600], sp('shapeType', 3)) +
          // Geometries the reader does not know: written as rectangles with their fill and outline.
          shp([0, 0, 1500, 600], sp('shapeType', 13) + sp('fillColor', 255)) +
          shp([0, 0, 1500, 600], sp('shapeType', 0), 'Freeform') +
          // A drawing object of Word 95: nothing is known of its look.
          '{\\*\\do\\dobxcolumn\\dobypara\\dptxbx{\\dptxbxtext Old\\par}\\dpx0\\dpy0\\dpxsize500\\dpysize500}',
      ),
    ),
  );
  assert.deepEqual(converted!.warnings, [
    '1 shape(s) were written as plain rectangles: their outline, fill and geometry were not carried over',
    '2 shape(s) were written as rectangles: their geometry was not carried over',
  ]);
  const xml = await documentXml(converted!.bytes);
  assert.ok(xml.includes('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="9525">'));
});

test('shapes converted to docx are drawn as they were: an RTF file and a doc file against the docx made of them', async () => {
  const doc = shapesDoc([
    { type: 20, rect: [0, 300, 3000, 300], props: [[P_LINE_COLOR, 255], [P_LINE_WIDTH, 19050]] },
    { type: 202, rect: [0, 600, 1500, 1200], text: 'Boxed', props: [[P_FILL_COLOR, 0x00ffff], [P_LINE_COLOR, 0xff0000], [P_LINE_WIDTH, 12700]] },
    { type: 3, rect: [1800, 600, 3000, 1200], props: [[P_FILL_FLAGS, NOT_FILLED], [P_LINE_DASHING, 6]] },
    { type: 2, rect: [0, 1500, 1500, 2100], props: [[P_LINE_FLAGS, NOT_LINED], [P_FILL_COLOR, 0x00ff00], [P_ADJUST, 2000]] },
    { type: 32, flags: 0x40, rect: [1800, 1500, 3000, 2100], props: [[P_LINE_WIDTH, 0]] },
    { type: 1, rect: [0, 2400, 1500, 3000], props: [[P_ROTATION, 30 << 16]] },
    { type: 3, rect: [1800, 2400, 3000, 3000], text: 'In', props: [[P_LINE_DASHING, 1]] },
  ]).build();
  for (const bytes of [rtf(RTF_SHAPES), doc]) {
    const file = await open(bytes);
    const before = painted(pageOf(await file.document!()));
    const converted = await convertForEditing(file);
    assert.deepEqual(converted!.warnings, []);
    const again = await open(converted!.bytes);
    assert.equal(again.format, 'docx');
    const after = painted(pageOf(await again.document!()));
    assert.ok(before.some((item) => item.includes('"type":"shape"')));
    assert.deepEqual(after, before);
  }
});
