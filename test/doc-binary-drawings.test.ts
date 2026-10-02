// Pictures and drawing objects of a Word binary file: inline pictures in the
// Data stream, floating shapes with their place, picture and text box.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { deflateRaw } from '../src/deflate';
import { DocLayout } from '../src/layout/doc/layout';
import type { ImageItem } from '../src/layout/doc/types';
import { paragraphText, type DrawingItem, type Paragraph } from '../src/model/doc/types';
import { fakeMeasurer, texts } from './doc-layout-fixtures';
import { DocBuilder, FC, cat, paragraphs, plc, runsOf, sprm, textOf, u16le, u32le, u8s, utf16le } from './doc-binary-fixtures';

/** An OfficeArt record: version, instance, type and content. */
const art = (version: number, instance: number, type: number, ...content: Uint8Array[]): Uint8Array => {
  const body = cat(...content);
  return cat(u16le(version | (instance << 4)), u16le(type), u32le(body.length), body);
};

/** A property table of a shape: simple properties as [id, value], text ones as [id, text]. */
function options(properties: [number, number | string][], type = 0xf00b): Uint8Array {
  const table: Uint8Array[] = [];
  const complex: Uint8Array[] = [];
  for (const [id, value] of properties) {
    if (typeof value === 'number') table.push(cat(u16le(id), u32le(value)));
    else {
      const text = cat(utf16le(value), u16le(0));
      table.push(cat(u16le(id | 0x8000), u32le(text.length)));
      complex.push(text);
    }
  }
  return art(3, properties.length, type, ...table, ...complex);
}

/** A shape container: its id and flags, then its property tables. */
const shape = (id: number, flags: number, ...tables: Uint8Array[]): Uint8Array => art(0xf, 0, 0xf004, art(2, 75, 0xf00a, u32le(id), u32le(flags)), ...tables);

const PNG = cat([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], new Uint8Array(40).fill(7));
const JPEG = cat([0xff, 0xd8, 0xff, 0xe0], new Uint8Array(30).fill(9));

/** A picture record of a bitmap kind: one identifier, a tag byte, the data. */
const bitmapBlip = (type: number, instance: number, data: Uint8Array): Uint8Array => art(0, instance, type, new Uint8Array(16), u8s(0xff), data);
const pngBlip = (data = PNG) => bitmapBlip(0xf01e, 0x6e0, data);

/** An entry of a picture store with the picture inside it, or one that points into the WordDocument stream. */
function storeEntry(blip: Uint8Array | undefined, delay = 0xffffffff, size = 0): Uint8Array {
  const fbse = new Uint8Array(36);
  fbse.set(u32le(blip ? blip.length : size), 20);
  fbse.set(u32le(1), 24);
  fbse.set(u32le(delay), 28);
  return art(2, 6, 0xf007, fbse, blip || new Uint8Array(0));
}

interface PictureSpec {
  width?: number;
  height?: number;
  scale?: [number, number];
  crop?: [number, number, number, number];
  mode?: number;
  name?: string;
}

/** A PICF with what follows it; returns the bytes for the Data stream. */
function picf(spec: PictureSpec, ...records: Uint8Array[]): Uint8Array {
  const header = new Uint8Array(68);
  const view = new DataView(header.buffer);
  const mode = spec.mode === undefined ? 0x64 : spec.mode;
  const name = mode === 0x66 ? cat(u8s((spec.name || '').length), [...(spec.name || '')].map((c) => c.charCodeAt(0))) : new Uint8Array(0);
  const body = cat(name, ...records);
  view.setUint32(0, 68 + body.length, true);
  view.setUint16(4, 68, true);
  view.setUint16(6, mode, true);
  view.setUint16(28, spec.width === undefined ? 1440 : spec.width, true);
  view.setUint16(30, spec.height === undefined ? 720 : spec.height, true);
  view.setUint16(32, spec.scale ? spec.scale[0] : 1000, true);
  view.setUint16(34, spec.scale ? spec.scale[1] : 1000, true);
  (spec.crop || [0, 0, 0, 0]).forEach((value, i) => view.setInt16(36 + i * 2, value, true));
  return cat(header, body);
}

/** A document "A" + picture character + "B" whose picture is at `offset` of the Data stream. */
function inlineDoc(data: Uint8Array, ...extra: Uint8Array[]): DocBuilder {
  const builder = new DocBuilder();
  const text = 'A\u0001B\r';
  builder.text(text);
  builder.papxAll(text);
  // Something else in front, so that the offset is not zero.
  builder.dataBlob(new Uint8Array(33));
  const at = builder.dataBlob(data);
  builder.chpx(1, 2, sprm(0x0855, 1), sprm(0x6a03, u32le(at)), ...extra);
  return builder;
}

const drawingsOf = (paragraph: Paragraph): DrawingItem[] =>
  runsOf(paragraph)
    .flatMap((run) => run.content)
    .filter((item): item is DrawingItem => item.type === 'drawing');

test('an inline picture: size from the PICF, the picture from the store entry behind the shape', () => {
  const data = picf({}, shape(1024, 0xa00, options([[0x0104, 1], [0x0381, 'A description']])), storeEntry(pngBlip()));
  const model = inlineDoc(data).read();
  const [paragraph] = paragraphs(model.body.blocks);
  assert.equal(paragraphText(paragraph), 'AB');
  const [drawing] = drawingsOf(paragraph);
  assert.equal(drawing.placement, 'inline');
  assert.equal(drawing.kind, 'picture');
  // One inch by half an inch, in EMU.
  assert.equal(drawing.width, 914400);
  assert.equal(drawing.height, 457200);
  assert.equal(drawing.description, 'A description');
  assert.deepEqual(drawing.image, { rId: 'doc:image1', target: '/doc/media/image1.png' });
  assert.deepEqual(drawing.images, [{ image: drawing.image }]);
  assert.equal(drawing.crop, undefined);
  const media = model.media!.get('/doc/media/image1.png')!;
  assert.equal(media.type, 'image/png');
  assert.deepEqual([...media.bytes], [...PNG]);
});

test('scaling and cropping of an inline picture; a picture that is a record of its own; the name of a linked file', () => {
  const data = picf({ width: 2000, height: 1000, scale: [500, 2000], crop: [200, 100, 300, 0], mode: 0x66, name: 'photo.jpg' }, shape(1024, 0xa00, options([[0x0104, 1]])), bitmapBlip(0xf01d, 0x46a, JPEG));
  const model = inlineDoc(data).read();
  const [drawing] = drawingsOf(paragraphs(model.body.blocks)[0]);
  // (2000 - 200 - 300) * 0.5 twips wide, (1000 - 100) * 2 twips high.
  assert.equal(drawing.width, 750 * 635);
  assert.equal(drawing.height, 1800 * 635);
  assert.deepEqual(drawing.crop, { left: 10000, right: 15000, top: 10000 });
  assert.equal(drawing.name, 'photo.jpg');
  const media = model.media!.get(drawing.image!.target!)!;
  assert.equal(media.type, 'image/jpeg');
  assert.equal(drawing.image!.target, '/doc/media/image1.jpg');
  assert.deepEqual([...media.bytes], [...JPEG]);
});

test('a device independent bitmap is handed out as a BMP file', () => {
  const dib = new Uint8Array(40 + 8 + 16);
  const view = new DataView(dib.buffer);
  view.setUint32(0, 40, true);
  view.setInt32(4, 4, true);
  view.setInt32(8, 4, true);
  view.setUint16(12, 1, true);
  view.setUint16(14, 1, true); // one bit a pixel: a colour table of two entries
  const data = picf({}, shape(1, 0xa00, options([[0x0104, 1]])), storeEntry(bitmapBlip(0xf01f, 0x7a8, dib)));
  const model = inlineDoc(data).read();
  const media = model.media!.get('/doc/media/image1.bmp')!;
  assert.equal(media.type, 'image/bmp');
  assert.equal(media.bytes.length, 14 + dib.length);
  assert.deepEqual([...media.bytes.subarray(0, 2)], [0x42, 0x4d]);
  const header = new DataView(media.bytes.buffer, media.bytes.byteOffset);
  assert.equal(header.getUint32(2, true), media.bytes.length);
  assert.equal(header.getUint32(10, true), 14 + 40 + 8);
  assert.deepEqual([...media.bytes.subarray(14)], [...dib]);
});

/** A metafile picture record: sizes, bounds, how it is packed, the data. */
function metafileBlip(packed: Uint8Array, size: number, compression: number): Uint8Array {
  const header = new Uint8Array(34);
  const view = new DataView(header.buffer);
  view.setUint32(0, size, true);
  view.setUint32(28, packed.length, true);
  header[32] = compression;
  header[33] = 0xfe;
  return art(0, 0x3d4, 0xf01a, new Uint8Array(16), header, packed);
}

test('a metafile is unpacked; one that lies about its size is not', () => {
  const emf = new Uint8Array(3000);
  for (let i = 0; i < emf.length; i++) emf[i] = (i * 7) & 0xff;
  const zlib = cat([0x78, 0x9c], deflateRaw(emf), [0, 0, 0, 0]);
  const read = (blip: Uint8Array) => {
    const model = inlineDoc(picf({}, shape(1, 0xa00, options([[0x0104, 1]])), storeEntry(blip))).read();
    const [drawing] = drawingsOf(paragraphs(model.body.blocks)[0]);
    return { model, drawing };
  };
  const good = read(metafileBlip(zlib, emf.length, 0));
  assert.equal(good.drawing.image!.target, '/doc/media/image1.emf');
  const media = good.model.media!.get('/doc/media/image1.emf')!;
  assert.equal(media.type, 'image/x-emf');
  assert.deepEqual([...media.bytes], [...emf]);

  // Not packed at all.
  const plain = read(metafileBlip(emf, emf.length, 0xfe));
  assert.deepEqual([...plain.model.media!.get('/doc/media/image1.emf')!.bytes], [...emf]);

  // Says it unpacks to 100 bytes but holds 3000: the picture is left out, the drawing keeps its place and size.
  const short = read(metafileBlip(zlib, 100, 0));
  assert.equal(short.drawing.image, undefined);
  assert.equal(short.drawing.width, 914400);
  assert.equal(short.model.media, undefined);
  // Says it unpacks to four gigabytes.
  assert.equal(read(metafileBlip(zlib, 0xffffffff, 0)).drawing.image, undefined);
  // Packed data that is not a deflate stream.
  assert.equal(read(metafileBlip(new Uint8Array(50).fill(0xff), 3000, 0)).drawing.image, undefined);
});

test('the same picture used twice is one entry of the media', () => {
  const builder = new DocBuilder();
  const text = '\u0001\u0001\r';
  builder.text(text);
  builder.papxAll(text);
  const at = builder.dataBlob(picf({}, shape(1, 0xa00, options([[0x0104, 1]])), storeEntry(pngBlip())));
  builder.chpx(0, 1, sprm(0x0855, 1), sprm(0x6a03, u32le(at)));
  builder.chpx(1, 2, sprm(0x0855, 1), sprm(0x6a03, u32le(at)), sprm(0x0835, 1));
  const model = builder.read();
  const drawings = drawingsOf(paragraphs(model.body.blocks)[0]);
  assert.equal(drawings.length, 2);
  assert.equal(drawings[0].image, drawings[1].image);
  assert.equal(model.media!.size, 1);
});

test('a picture character that stands for field data or an embedded object is not a picture', () => {
  const data = picf({}, shape(1, 0xa00, options([[0x0104, 1]])), storeEntry(pngBlip()));
  for (const flag of [0x0806, 0x080a, 0x0856]) {
    const model = inlineDoc(data, sprm(flag, 1)).read();
    assert.deepEqual(drawingsOf(paragraphs(model.body.blocks)[0]), []);
    assert.deepEqual(textOf(model.body.blocks), ['AB']);
    assert.equal(model.media, undefined);
  }
});

test('a picture whose place is outside the Data stream, or holds no picture, is passed over and counted', () => {
  const builder = new DocBuilder();
  const text = 'A\u0001B\r';
  builder.text(text);
  builder.papxAll(text);
  builder.chpx(1, 2, sprm(0x0855, 1), sprm(0x6a03, u32le(5000)));
  const model = builder.read();
  assert.deepEqual(textOf(model.body.blocks), ['AB']);
  assert.equal(model.unknown.get('picture/unreadable'), 1);

  // A PICF whose length reaches past the stream.
  const cut = picf({}, shape(1, 0xa00, options([[0x0104, 1]])), storeEntry(pngBlip()));
  new DataView(cut.buffer).setUint32(0, 100000, true);
  assert.equal(inlineDoc(cut).read().unknown.get('picture/unreadable'), 1);

  // A shape without a picture behind it: the drawing is there, without an image.
  const empty = inlineDoc(picf({}, shape(1, 0xa00, options([[0x0104, 3]])))).read();
  const [drawing] = drawingsOf(paragraphs(empty.body.blocks)[0]);
  assert.equal(drawing.image, undefined);
  assert.equal(drawing.kind, 'picture');
});

test('records that claim to be longer than what holds them end the reading there', () => {
  const blip = pngBlip();
  // The picture record says it is far longer than the store entry around it.
  new DataView(blip.buffer).setUint32(4, 0x7fffffff, true);
  const model = inlineDoc(picf({}, shape(1, 0xa00, options([[0x0104, 1]])), storeEntry(blip))).read();
  const [drawing] = drawingsOf(paragraphs(model.body.blocks)[0]);
  assert.equal(drawing.image, undefined);
  // A property table that counts more properties than it has room for.
  const table = options([[0x0104, 1]]);
  new DataView(table.buffer).setUint16(0, 3 | (4000 << 4), true);
  const other = inlineDoc(picf({}, shape(1, 0xa00, table), storeEntry(pngBlip()))).read();
  assert.equal(drawingsOf(paragraphs(other.body.blocks)[0]).length, 1);
});

/**
 * A document with two drawing objects in its second paragraph: a picture and
 * a text box. The picture of the first lies in the WordDocument stream, where
 * the picture store of the drawing points to.
 */
function floatingDoc(pictureFlags = 0, pictureProps: [number, number | string][] = [], textProps: [number, number | string][] = []): DocBuilder {
  const builder = new DocBuilder();
  const body = 'Before\rAn\u0008chor\u0008ed\rAfter\r';
  const box = 'In the box\rsecond\r\r';
  const text = body + box + '\r';
  builder.text(text);
  builder.papxAll(text);
  builder.chpx(9, 10, sprm(0x0855, 1)).chpx(14, 15, sprm(0x0855, 1));
  builder.ccp.txbx = box.length;
  const blipAt = builder.blob(pngBlip());

  const spa = (id: number, rect: number[], flags: number) => cat(u32le(id), ...rect.map(u32le), u16le(flags), u32le(0));
  builder.table(
    FC.PlcSpaMom,
    plc(
      [9, 14, body.length],
      [
        // 1 inch from the page edge, 2 inches below the margin; 3 by 1.5 inches; wrapped around on both sides.
        spa(1026, [1440, 2880, 1440 + 4320, 2880 + 2160], (1 << 1) | (0 << 3) | (2 << 5) | pictureFlags),
        // Relative to the column and the paragraph, no wrapping.
        spa(1027, [100, 200, 3100, 1200], (2 << 1) | (2 << 3) | (3 << 5)),
      ],
    ),
  );
  const patriarch = shape(1024, 0x05);
  const picture = shape(1026, 0xa00, options([[0x0104, 1], [0x0380, 'Picture 1'], ...pictureProps]));
  const textBox = shape(1027, 0xa00, options([[0x0080, 1 << 16], [0x0081, 91440], [0x0082, 45720], [0x0083, 91440], [0x0084, 45720], ...textProps]));
  const store = art(0xf, 1, 0xf001, storeEntry(undefined, blipAt, pngBlip().length));
  builder.table(
    FC.DggInfo,
    cat(art(0xf, 0, 0xf000, art(0, 0, 0xf006, new Uint8Array(16)), store), u8s(0), art(0xf, 0, 0xf002, art(0, 1, 0xf008, new Uint8Array(8)), art(0xf, 0, 0xf003, patriarch, picture, textBox))),
  );
  const entry = (shapeId: number) => {
    const bytes = new Uint8Array(22);
    bytes.set(u32le(shapeId), 14);
    return bytes;
  };
  builder.table(FC.PlcftxbxTxt, plc([0, box.length - 1, box.length], [entry(1027), entry(0)]));
  return builder;
}

test('drawing objects: place, wrapping, the picture from the store, the text of a text box', () => {
  const model = floatingDoc().read();
  assert.deepEqual(textOf(model.body.blocks), ['Before', 'Anchored', 'After']);
  const [picture, box] = drawingsOf(paragraphs(model.body.blocks)[1]);
  assert.equal(picture.placement, 'anchor');
  assert.equal(picture.kind, 'picture');
  assert.equal(picture.name, 'Picture 1');
  assert.equal(picture.width, 4320 * 635);
  assert.equal(picture.height, 2160 * 635);
  assert.deepEqual(picture.anchor, {
    horizontal: { relativeFrom: 'page', offset: 1440 * 635 },
    vertical: { relativeFrom: 'margin', offset: 2880 * 635 },
    wrap: 'square',
    wrapText: 'bothSides',
  });
  assert.equal(picture.image!.target, '/doc/media/image1.png');
  assert.equal(model.media!.get('/doc/media/image1.png')!.bytes.length, 48);

  assert.equal(box.kind, 'shape');
  assert.equal(box.image, undefined);
  assert.deepEqual(box.anchor, { horizontal: { relativeFrom: 'column', offset: 100 * 635 }, vertical: { relativeFrom: 'paragraph', offset: 200 * 635 }, wrap: 'none' });
  assert.equal(box.textBoxes!.length, 1);
  assert.deepEqual(textOf(box.textBoxes![0]), ['In the box', 'second']);
});

test('what the shape itself says about its place wins over the rectangle; behind the text, rotated, flipped', () => {
  const model = floatingDoc(
    0x4000,
    [
      [0x038f, 2], // centred
      [0x0390, 0], // between the margins
      [0x0391, 3], // at the bottom
      [0x0392, 1], // of the page
      [0x0384, 114300],
      [0x0386, 228600],
      [0x0004, 90 << 16],
      [0x0100, 0x4000], // a quarter cut off the top
    ],
    [[0x03bf, 0x00200000]],
  ).read({
    patch: ({ table, fcLcb }) => {
      // Flip the picture: the flags of its shape are 8 bytes into the record of the shape.
      const [at, size] = fcLcb(FC.DggInfo);
      for (let p = at; p + 16 <= at + size; p++) {
        const view = new DataView(table.buffer, table.byteOffset);
        if (view.getUint16(p + 2, true) === 0xf00a && view.getUint32(p + 8, true) === 1026) view.setUint32(p + 12, 0xa00 | 0x40, true);
      }
    },
  });
  const [picture, box] = drawingsOf(paragraphs(model.body.blocks)[1]);
  assert.deepEqual(picture.anchor!.horizontal, { relativeFrom: 'margin', align: 'center' });
  assert.deepEqual(picture.anchor!.vertical, { relativeFrom: 'page', align: 'bottom' });
  assert.equal(picture.anchor!.behindDoc, true);
  assert.equal(picture.anchor!.distL, 114300);
  assert.equal(picture.anchor!.distR, 228600);
  assert.deepEqual(picture.transform, { rotation: 90 * 60000, flipH: true });
  assert.deepEqual(picture.crop, { top: 25000 });
  // The text box says "not behind the text" itself.
  assert.equal(box.anchor!.behindDoc, undefined);
});

test('the layout engine places the drawing objects of a doc file and lays out the text of the box', () => {
  const model = floatingDoc().read();
  const layout = new DocLayout(model, fakeMeasurer);
  assert.equal(layout.layout().done, true);
  const page = layout.pages[0];
  const images = page.items.filter((item): item is ImageItem => item.type === 'image');
  assert.equal(images.length, 1);
  // The part name of the picture is the key into DocModel.media.
  assert.equal(images[0].part, '/doc/media/image1.png');
  assert.equal(images[0].x, 96);
  assert.equal(images[0].width, 288);
  assert.equal(images[0].height, 144);
  const shown = texts(page).map((item) => item.text);
  assert.ok(shown.includes('In the box'), shown.join('|'));
  assert.ok(shown.includes('second'));
  // The insets of the box (0.1 inch) are where layout looks for them: the text starts 9.6 px inside the box.
  const boxText = texts(page).find((item) => item.text === 'In the box')!;
  const anchor = texts(page).find((item) => item.text === 'Before')!;
  assert.ok(Math.abs(boxText.x - (anchor.x + 100 / 15 + 9.6)) < 0.01, `${boxText.x} ${anchor.x}`);
});

test('drawing objects in a header have their own list and their own text boxes', () => {
  const builder = new DocBuilder();
  const body = 'Body\r';
  // Header stories: six empty ones for the notes, then the odd header of the only section.
  const header = 'H\u0008\r\r';
  const box = 'Header box\r\r';
  const text = body + header + box + '\r';
  builder.text(text);
  builder.papxAll(text);
  builder.chpx(body.length + 1, body.length + 2, sprm(0x0855, 1));
  builder.ccp.hdd = header.length;
  builder.ccp.hdrTxbx = box.length;
  builder.table(FC.PlcfHdd, plc([0, 0, 0, 0, 0, 0, 0, 0, 3, 3, 3, 3, 3, 4]));
  // Positions of header shapes count from the start of the header text.
  builder.table(FC.PlcSpaHdr, plc([1, 4], [cat(u32le(2050), u32le(0), u32le(0), u32le(2000), u32le(1000), u16le(0x0001 | (3 << 5)), u32le(0))]));
  builder.table(FC.DggInfo, cat(art(0xf, 0, 0xf000), u8s(1), art(0xf, 0, 0xf002, art(0xf, 0, 0xf003, shape(2048, 0x05), shape(2050, 0xa00, options([[0x0080, 1 << 16]]))))));
  const entry = new Uint8Array(22);
  entry.set(u32le(2050), 14);
  builder.table(FC.PlcfHdrtxbxTxt, plc([0, box.length - 1, box.length], [entry, new Uint8Array(22)]));
  const model = builder.read();
  const story = model.headers.get('hdr:0:default')!;
  assert.deepEqual(textOf(story.blocks), ['H']);
  const [drawing] = drawingsOf(paragraphs(story.blocks)[0]);
  assert.deepEqual(textOf(drawing.textBoxes![0]), ['Header box']);
  assert.deepEqual(textOf(model.body.blocks), ['Body']);
});

test('a drawing whose shape is missing, or whose drawing data is damaged, keeps its place', () => {
  const missing = floatingDoc().read({
    patch: ({ table, fcLcb }) => {
      const [at, size] = fcLcb(FC.DggInfo);
      table.fill(0xff, at, at + size);
    },
  });
  const [picture, box] = drawingsOf(paragraphs(missing.body.blocks)[1]);
  assert.equal(picture.kind, 'shape');
  assert.equal(picture.image, undefined);
  assert.equal(picture.width, 4320 * 635);
  // The text box is still found through the list of text boxes, which names its shape.
  assert.deepEqual(textOf(box.textBoxes![0]), ['In the box', 'second']);
  assert.deepEqual(textOf(missing.body.blocks), ['Before', 'Anchored', 'After']);

  // The picture store points outside the WordDocument stream.
  const lost = floatingDoc().read({
    patch: ({ table, fcLcb }) => {
      const [at, size] = fcLcb(FC.DggInfo);
      const view = new DataView(table.buffer, table.byteOffset);
      for (let p = at; p + 44 <= at + size; p++) if (view.getUint16(p + 2, true) === 0xf007) view.setUint32(p + 8 + 28, 0x7ffffff0, true);
    },
  });
  assert.equal(drawingsOf(paragraphs(lost.body.blocks)[1])[0].image, undefined);
  assert.equal(lost.media, undefined);
});

test('groups nested without end do not send the reader down without end', () => {
  // A group container in a group container, two thousand deep.
  let group = art(0xf, 0, 0xf003, shape(5000, 0x01));
  for (let i = 0; i < 2000; i++) group = art(0xf, 0, 0xf003, shape(5001 + i, 0x01), group);
  const builder = new DocBuilder();
  const text = 'A\u0008\r';
  builder.text(text);
  builder.papxAll(text);
  builder.chpx(1, 2, sprm(0x0855, 1));
  // The drawing object is the outermost group.
  builder.table(FC.PlcSpaMom, plc([1, 3], [cat(u32le(7000), u32le(0), u32le(0), u32le(100), u32le(100), u16le(0), u32le(0))]));
  builder.table(FC.DggInfo, cat(art(0xf, 0, 0xf000), u8s(0), art(0xf, 0, 0xf002, group)));
  const model = builder.read();
  const [drawing] = drawingsOf(paragraphs(model.body.blocks)[0]);
  assert.equal(drawing.kind, 'group');
  assert.deepEqual(textOf(model.body.blocks), ['A']);
});
