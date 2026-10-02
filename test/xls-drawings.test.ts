// Pictures, shapes and comments of an xls workbook: the OfficeArt drawing of a
// sheet with its OBJ, TXO and NOTE records, and the picture store of the
// workbook. Every stream is built here, record by record.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { png } from '../e2e/fixtures/docx';
import { open } from '../src/index';
import { readXlsStream } from '../src/formats/xls/workbook';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import { getCell, type DrawingGroup, type DrawingPicture, type DrawingShape, type Workbook } from '../src/model/sheet/types';
import { sniffImageType } from '../src/render/images';
import { SheetDrawings, anchorRect } from '../src/render/sheet-drawings';
import { bof, cat, defaultStyles, eof, font, narrow, number, rec, str16, u16, u32, workbookStream, xf, xlsFile, type Bytes } from './xls-fixtures';

// --- OfficeArt ---------------------------------------------------------------------------

const artHeader = (type: number, instance: number, version: number, length: number): number[] => [...u16((instance << 4) | version), ...u16(type), ...u32(length)];
const art = (type: number, instance: number, version: number, content: Bytes = []): number[] => [...artHeader(type, instance, version, content.length), ...content];

/** A run of OfficeArt bytes and the BIFF records that stand after it in the sheet. */
interface Segment {
  art: number[];
  after: Uint8Array[];
}

const sizeOf = (segments: Segment[]): number => segments.reduce((sum, segment) => sum + segment.art.length, 0);

/** A container around segments: its header counts all the OfficeArt bytes, whatever records stand between them. */
function wrap(type: number, segments: Segment[]): Segment[] {
  const head = artHeader(type, 0, 0xf, sizeOf(segments));
  return [{ art: [...head, ...segments[0].art], after: segments[0].after }, ...segments.slice(1)];
}

type Prop = [id: number, value: number] | [id: number, text: string];

function options(props: Prop[]): number[] {
  const table: number[] = [];
  const complex: number[] = [];
  for (const [id, value] of props) {
    if (typeof value === 'string') {
      const data: number[] = [];
      for (let i = 0; i < value.length; i++) data.push(value.charCodeAt(i) & 0xff, value.charCodeAt(i) >> 8);
      data.push(0, 0);
      table.push(...u16(id | 0x8000), ...u32(data.length));
      complex.push(...data);
    } else table.push(...u16(id), ...u32(value));
  }
  return art(0xf00b, props.length, 3, [...table, ...complex]);
}

/** Flags, then column, offset (1/1024), row, offset (1/256) of the top left and the bottom right corner. */
const clientAnchor = (from: [number, number, number, number], to: [number, number, number, number], flags = 0): number[] =>
  art(0xf010, 0, 0, [...u16(flags), ...u16(from[0]), ...u16(from[1]), ...u16(from[2]), ...u16(from[3]), ...u16(to[0]), ...u16(to[1]), ...u16(to[2]), ...u16(to[3])]);

const childAnchor = (left: number, top: number, right: number, bottom: number): number[] => art(0xf00f, 0, 0, [...u32(left), ...u32(top), ...u32(right), ...u32(bottom)]);

const obj = (type: number, id: number): Uint8Array => rec(0x005d, u16(0x15), u16(0x12), u16(type), u16(id), u16(0x6011), new Array(12).fill(0), u16(0), u16(0));

/** TXO with its two CONTINUE records. `runs`: first character and font index of each run. */
function txo(text: string, runs: [number, number][] = [[0, 0]], flags = 0x0212): Uint8Array[] {
  const formatting = [...runs, [text.length, 0]].flatMap(([first, fontIndex]) => [...u16(first), ...u16(fontIndex), ...u32(0)]);
  return [rec(0x01b6, u16(flags), u16(0), new Array(6).fill(0), u16(text.length), u16(formatting.length), u32(0)), rec(0x003c, [0], narrow(text)), rec(0x003c, formatting)];
}

const note = (row: number, col: number, id: number, author: string, visible = false): Uint8Array => rec(0x001c, u16(row), u16(col), u16(visible ? 2 : 0), u16(id), str16(author), [0]);

interface ShapeSpec {
  /** MSOSPT: 75 picture frame, 1 rectangle, 3 ellipse, 20 line, 202 text box. */
  kind: number;
  /** Object type of the OBJ record: 8 picture, 2 rectangle, 6 text, 0x19 note, 5 chart, 0x14 drop-down. */
  ot: number;
  id: number;
  props?: Prop[];
  anchor?: number[];
  flags?: number;
  text?: Uint8Array[];
  /** Records after the OBJ record (the substream of a chart). */
  extra?: Uint8Array[];
}

function shape(spec: ShapeSpec): Segment[] {
  const body = [...art(0xf00a, spec.kind, 2, [...u32(1024 + spec.id), ...u32(spec.flags ?? 0xa00)]), ...options(spec.props ?? []), ...(spec.anchor ?? clientAnchor([1, 0, 1, 0], [3, 0, 5, 0])), ...art(0xf011, 0, 0)];
  const first: Segment = { art: body, after: [obj(spec.ot, spec.id), ...(spec.extra ?? [])] };
  const segments = spec.text ? [first, { art: art(0xf00d, 0, 0), after: spec.text }] : [first];
  return wrap(0xf004, segments);
}

/** A group: the shape that stands for it (with the coordinates of its members), then the members. */
function groupOf(id: number, space: [number, number, number, number], anchor: number[], members: Segment[][]): Segment[] {
  const self: Segment = {
    art: [...art(0xf009, 0, 1, [...u32(space[0]), ...u32(space[1]), ...u32(space[2]), ...u32(space[3])]), ...art(0xf00a, 0, 2, [...u32(1024 + id), ...u32(0x201)]), ...anchor, ...art(0xf011, 0, 0)],
    after: [obj(0, id)],
  };
  return wrap(0xf003, [...wrap(0xf004, [self]), ...members.flat()]);
}

/** The records of a sheet's drawing: the drawing container cut into MSODRAWING records with the object records between them. */
function drawingRecords(shapes: Segment[][]): Uint8Array {
  const patriarch: Segment = { art: art(0xf004, 0, 0xf, [...art(0xf009, 0, 1, new Array(16).fill(0)), ...art(0xf00a, 0, 2, [...u32(1024), ...u32(5)])]), after: [] };
  const group = wrap(0xf003, [patriarch, ...shapes.flat()]);
  // The patriarch has no object record: its bytes go into the same record as the first shape.
  const merged = group.length > 1 ? [{ art: [...group[0].art, ...group[1].art], after: group[1].after }, ...group.slice(2)] : group;
  const dg = wrap(0xf002, [{ art: art(0xf008, 1, 0, [...u32(shapes.length + 1), ...u32(1024 + shapes.length)]), after: [] }, ...merged]);
  const all = dg.length > 1 ? [{ art: [...dg[0].art, ...dg[1].art], after: dg[1].after }, ...dg.slice(2)] : dg;
  return cat(...all.map((segment) => cat(rec(0x00ec, segment.art), ...segment.after)));
}

/** A BLIP inside its store entry: a PNG (type 6) unless stated. */
function bse(data: Uint8Array, blipType = 0xf01e, kind = 6): number[] {
  const blip = blipType === 0xf01a ? art(blipType, 0x3d4, 0, [...new Array(16).fill(1), ...new Array(34).fill(0), ...data]) : art(blipType, 0x6e0, 0, [...new Array(16).fill(1), 0xff, ...data]);
  return art(0xf007, kind, 2, [kind, kind, ...new Array(16).fill(1), ...u16(0xff), ...u32(blip.length), ...u32(1), ...u32(0), 0, 0, 0, 0, ...blip]);
}

/** The drawing group of the globals with its picture store, as one run of OfficeArt bytes. */
function drawingGroup(entries: number[][]): number[] {
  const store = art(0xf001, entries.length, 0xf, entries.flat());
  return art(0xf000, 0, 0xf, [...art(0xf006, 0, 0, new Array(24).fill(0)), ...store]);
}

const RED: [number, number, number] = [220, 0, 0];

function book(sheetRecords: Bytes, tail: Bytes = [], head?: Bytes): Workbook {
  return readXlsStream(workbookStream({ head, tail, sheets: [{ name: 'S', records: sheetRecords }] })).workbook;
}

// --- pictures ----------------------------------------------------------------------------

test('xls drawings: a picture is read with its anchor, and its bytes come from the store when asked for', async () => {
  const image = png(8, 4, RED);
  const records = drawingRecords([shape({ kind: 75, ot: 8, id: 1, props: [[0x4104, 1], [0x0380, 'Logo'], [0x0381, 'Company logo']], anchor: clientAnchor([1, 512, 2, 64], [3, 256, 6, 128]) })]);
  const workbook = book(cat(number(0, 0, 1), records, number(1, 0, 2)), rec(0x00eb, drawingGroup([bse(image)])));
  const sheet = workbook.sheets[0];
  assert.equal(sheet.drawings?.length, 1);
  assert.deepEqual(sheet.drawings?.[0].anchor, {
    kind: 'twoCell',
    from: { col: 1, colOff: 0, colShare: 0.5, row: 2, rowOff: 0, rowShare: 0.25 },
    to: { col: 3, colOff: 0, colShare: 0.25, row: 6, rowOff: 0, rowShare: 0.5 },
  });
  const picture = sheet.drawings?.[0].object as DrawingPicture;
  assert.deepEqual(picture, { name: 'Logo', description: 'Company logo', type: 'picture', media: 'xls:blip:1' });
  assert.ok(workbook.media);
  const bytes = await workbook.media.readPart('xls:blip:1');
  assert.deepEqual(bytes, image);
  assert.equal(sniffImageType(bytes), 'image/png');
  // Only pictures a drawing shows can be asked for.
  await assert.rejects(workbook.media.readPart('xls:blip:2'));
  await assert.rejects(workbook.media.readPart('/xl/workbook.xml'));
  // The cells around the drawing records are read as before.
  assert.equal(getCell(sheet, 0, 0)?.value, 1);
  assert.equal(getCell(sheet, 1, 0)?.value, 2);
  assert.equal(sheet.warnings, undefined);
  // In pixels: half of column B (64 px) in, a quarter of row 3 (20 px) down.
  assert.deepEqual(anchorRect(sheet.drawings![0].anchor, createSheetMetrics(sheet)), { x: 64 + 32, y: 40 + 5, width: 3 * 64 + 16 - 96, height: 6 * 20 + 10 - 45 });
});

test('xls drawings: the flags of the anchor say what inserted rows do to the picture', () => {
  const records = drawingRecords([
    shape({ kind: 75, ot: 8, id: 1, anchor: clientAnchor([0, 0, 0, 0], [1, 0, 1, 0], 0) }),
    shape({ kind: 75, ot: 8, id: 2, anchor: clientAnchor([0, 0, 0, 0], [1, 0, 1, 0], 2) }),
    shape({ kind: 75, ot: 8, id: 3, anchor: clientAnchor([0, 0, 0, 0], [1, 0, 1, 0], 3) }),
  ]);
  assert.deepEqual(
    book(records).sheets[0].drawings?.map((drawing) => drawing.anchor.editAs),
    [undefined, 'oneCell', 'absolute'],
  );
});

test('xls drawings: the store may be cut into several records, and a drawing into CONTINUE records', async () => {
  const image = png(64, 64, RED, [0, 0, 220]);
  const group = drawingGroup([bse(png(2, 2, RED)), bse(image)]);
  const third = Math.floor(group.length / 3);
  const tail = cat(rec(0x00eb, group.slice(0, third)), rec(0x003c, group.slice(third, 2 * third)), rec(0x00eb, group.slice(2 * third)));
  const whole = drawingRecords([shape({ kind: 75, ot: 8, id: 1, props: [[0x4104, 2]] })]);
  // The same drawing with its first record cut in two: the second half in a CONTINUE record.
  const firstLength = whole[2] | (whole[3] << 8);
  const cut = cat(rec(0x00ec, whole.subarray(4, 24)), rec(0x003c, whole.subarray(24, 4 + firstLength)), whole.subarray(4 + firstLength));
  for (const records of [whole, cut]) {
    const workbook = book(records, tail);
    assert.equal((workbook.sheets[0].drawings?.[0].object as DrawingPicture).media, 'xls:blip:2');
    assert.deepEqual(await workbook.media?.readPart('xls:blip:2'), image);
  }
});

test('xls drawings: crop, rotation and flips', () => {
  const records = drawingRecords([
    shape({ kind: 75, ot: 8, id: 1, flags: 0xa00 | 0x40 | 0x80, props: [[0x4104, 1], [0x0100, 0x2000], [0x0101, 0x4000], [0x0102, 0x8000], [0x0103, 0xffff0000], [0x0004, 90 * 65536]] }),
    shape({ kind: 75, ot: 8, id: 2, props: [[0x4104, 1], [0x0102, 0xc000], [0x0103, 0xc000], [0x0004, (-90 * 65536) >>> 0]] }),
  ]);
  const drawings = book(records, rec(0x00eb, drawingGroup([bse(png(2, 2, RED))]))).sheets[0].drawings;
  const first = drawings?.[0].object as DrawingPicture;
  // 16.16 fixed point shares; a negative one (room added around the picture) is not shown.
  assert.deepEqual(first.crop, { left: 0.5, top: 0.125, right: 0, bottom: 0.25 });
  assert.equal(first.rotation, 90);
  assert.equal(first.flipH, true);
  assert.equal(first.flipV, true);
  const second = drawings?.[1].object as DrawingPicture;
  // A crop that leaves nothing is not applied.
  assert.equal(second.crop, undefined);
  assert.equal(second.rotation, 270);
});

test('xls drawings: a picture whose store entry is missing, empty or a metafile is a placeholder', async () => {
  const emf = new Uint8Array(100);
  const group = drawingGroup([art(0xf007, 2, 2, new Array(36).fill(0)), bse(emf, 0xf01a, 2), bse(png(2, 2, RED))]);
  const records = drawingRecords([
    shape({ kind: 75, ot: 8, id: 1, props: [[0x4104, 1]] }),
    shape({ kind: 75, ot: 8, id: 2, props: [[0x4104, 2]] }),
    shape({ kind: 75, ot: 8, id: 3, props: [[0x4104, 9]] }),
    shape({ kind: 75, ot: 8, id: 4 }),
    shape({ kind: 75, ot: 8, id: 5, props: [[0x4104, 3]] }),
  ]);
  const workbook = book(records, rec(0x00eb, group));
  const media = workbook.sheets[0].drawings?.map((drawing) => (drawing.object as DrawingPicture).media);
  assert.deepEqual(media, [undefined, 'xls:blip:2', undefined, undefined, 'xls:blip:3']);
  // The metafile has an entry, but it is never unpacked or handed out.
  await assert.rejects(workbook.media!.readPart('xls:blip:2'), /not of a kind that is shown/);
  assert.equal(sniffImageType(await workbook.media!.readPart('xls:blip:3')), 'image/png');
});

// --- shapes ------------------------------------------------------------------------------

test('xls drawings: shapes with fill, outline and text', () => {
  const head = cat(defaultStyles(), font({ bold: true, name: 'Arial', size: 14 }));
  const records = drawingRecords([
    // Fill and line as red, green, blue in the low bytes; a line of two pixels, dashed.
    shape({ kind: 1, ot: 2, id: 1, props: [[0x0181, 0x0050b000], [0x01c0, 0x00332211], [0x01cb, 19050], [0x01ce, 1]] }),
    // Colours of the palette (10 is red), and the two system colours shapes usually have.
    shape({ kind: 3, ot: 3, id: 2, props: [[0x0181, 0x0800000a], [0x01c0, 0x08000040]] }),
    // Nothing stated: white with a black hairline.
    shape({ kind: 202, ot: 6, id: 3, text: txo('Merhaba kalin', [[0, 0], [8, 5]], 0x0224), props: [[0x0081, 0], [0x0082, 19050]] }),
    // Neither filled nor outlined.
    shape({ kind: 1, ot: 2, id: 4, props: [[0x01bf, 0x00100000], [0x01ff, 0x00080000]] }),
    shape({ kind: 20, ot: 1, id: 5, flags: 0xa00 | 0x40, props: [[0x01c0, 0x00ff0000]] }),
    // A shape type that has no outline of its own here: its box.
    shape({ kind: 187, ot: 0x1e, id: 6, props: [[0x0181, 0x10000005]] }),
  ]);
  const shapes = book(records, [], head).sheets[0].drawings?.map((drawing) => drawing.object as DrawingShape);
  assert.ok(shapes);
  assert.deepEqual(shapes[0], { type: 'shape', geometry: 'rect', fill: '#00B050', line: { color: '#112233', width: 2, dash: true } });
  assert.deepEqual(shapes[1], { type: 'shape', geometry: 'ellipse', fill: '#FF0000', line: { color: '#000000', width: 1 } });
  assert.equal(shapes[2].fill, '#FFFFFF');
  assert.deepEqual(shapes[2].line, { color: '#000000', width: 1 });
  assert.deepEqual(shapes[2].text, {
    paragraphs: [{ runs: [{ text: 'Merhaba ', font: 'Arial', size: 10 }, { text: 'kalin', font: 'Arial', size: 14, bold: true }], align: 'center' }],
    anchor: 'middle',
    insets: [0, 2, 9.6, 4.8],
  });
  assert.deepEqual(shapes[3], { type: 'shape', geometry: 'rect' });
  assert.deepEqual(shapes[4], { flipH: true, type: 'shape', geometry: 'line', connector: true, line: { color: '#0000FF', width: 1 } });
  // A colour of the system has no fixed value: the default is used.
  assert.deepEqual(shapes[5], { type: 'shape', geometry: 'rect', fill: '#FFFFFF', line: { color: '#000000', width: 1 } });
});

test('xls drawings: the members of a group are placed by the coordinates of the group', () => {
  const member = (id: number, anchor: number[], fill: number) => shape({ kind: 1, ot: 2, id, anchor, props: [[0x0181, fill]] });
  const records = drawingRecords([
    groupOf(1, [100, 200, 300, 600], clientAnchor([0, 0, 0, 0], [4, 0, 8, 0]), [
      member(2, childAnchor(100, 200, 200, 400), 0x0000ff),
      member(3, childAnchor(200, 400, 300, 600), 0xff0000),
      // No place in the group: left out.
      member(4, clientAnchor([0, 0, 0, 0], [1, 0, 1, 0]), 0x00ff00),
    ]),
    shape({ kind: 1, ot: 2, id: 5, props: [[0x0181, 0x123456]] }),
  ]);
  const drawings = book(records).sheets[0].drawings;
  assert.equal(drawings?.length, 2);
  const top = drawings?.[0].object as DrawingGroup;
  assert.equal(top.type, 'group');
  assert.deepEqual(
    top.children.map((child) => [(child as DrawingShape).fill, child.box]),
    [
      ['#FF0000', { x: 0, y: 0, width: 0.5, height: 0.5 }],
      ['#0000FF', { x: 0.5, y: 0.5, width: 0.5, height: 0.5 }],
    ],
  );
  // The object records were handed out in order: the shape after the group has its own.
  assert.equal((drawings?.[1].object as DrawingShape).fill, '#563412');
});

test('xls drawings: a chart is a placeholder, its substream is skipped; controls are not drawings', () => {
  const chart = [bof(0x0020), number(5, 5, 555), rec(0x00ec, [1, 2, 3]), eof()];
  const records = drawingRecords([
    shape({ kind: 201, ot: 5, id: 1, extra: chart, props: [[0x0380, 'Chart 1']] }),
    // The button of a filter, a check box.
    shape({ kind: 201, ot: 0x14, id: 2 }),
    shape({ kind: 201, ot: 0x0b, id: 3 }),
    shape({ kind: 1, ot: 2, id: 4 }),
  ]);
  const sheet = book(cat(records, number(1, 1, 7))).sheets[0];
  assert.deepEqual(
    sheet.drawings?.map((drawing) => [drawing.object.type, drawing.object.name]),
    [
      ['placeholder', 'Chart 1'],
      ['shape', undefined],
    ],
  );
  assert.equal(getCell(sheet, 5, 5), undefined);
  assert.equal(getCell(sheet, 1, 1)?.value, 7);
  assert.equal(sheet.warnings, undefined);
});

test('xls drawings: a hidden shape is marked hidden and takes no place in the view', () => {
  const records = drawingRecords([shape({ kind: 1, ot: 2, id: 1, props: [[0x03bf, 0x00020002]] }), shape({ kind: 1, ot: 2, id: 2, props: [[0x03bf, 0x00020000]] })]);
  const sheet = book(records).sheets[0];
  assert.deepEqual(
    sheet.drawings?.map((drawing) => drawing.object.hidden),
    [true, undefined],
  );
  assert.equal(new SheetDrawings(sheet).rects(createSheetMetrics(sheet)).length, 1);
});

// --- comments ----------------------------------------------------------------------------

test('xls comments: a note is read with its cell, author, text and box; its box is not a drawing', () => {
  const head = cat(defaultStyles(), font({ bold: true, name: 'Tahoma', size: 9 }));
  const records = cat(
    drawingRecords([
      shape({ kind: 202, ot: 0x19, id: 1, anchor: clientAnchor([2, 512, 0, 128], [4, 0, 4, 0]), text: txo('Ayse:\nKontrol edildi', [[0, 5], [5, 0]]) }),
      shape({ kind: 1, ot: 2, id: 2 }),
      shape({ kind: 202, ot: 0x19, id: 3, text: txo('ikinci') }),
    ]),
    note(1, 1, 1, 'Ayse'),
    note(4, 2, 3, 'Bora', true),
    // A note whose object is missing, one outside the sheet and one for a cell that has one already.
    note(6, 0, 77, ''),
    note(1, 300, 3, 'x'),
    note(1, 1, 3, 'x'),
  );
  const { workbook, report } = readXlsStream(workbookStream({ head, sheets: [{ name: 'S', records }] }));
  const sheet = workbook.sheets[0];
  assert.equal(sheet.drawings?.length, 1);
  assert.equal(sheet.drawings?.[0].object.type, 'shape');
  assert.equal(report.comments, 5);
  const comments = sheet.comments;
  assert.ok(comments);
  assert.equal(comments.length, 3);
  assert.equal(comments[0].row, 1);
  assert.equal(comments[0].col, 1);
  assert.equal(comments[0].author, 'Ayse');
  assert.equal(comments[0].text, 'Ayse:\nKontrol edildi');
  assert.deepEqual(
    comments[0].runs?.map((run) => [run.text, run.font?.bold]),
    [
      ['Ayse:', true],
      ['\nKontrol edildi', false],
    ],
  );
  assert.equal(comments[0].visible, undefined);
  assert.deepEqual(comments[0].box, {
    kind: 'twoCell',
    from: { col: 2, colOff: 0, colShare: 0.5, row: 0, rowOff: 0, rowShare: 0.5 },
    to: { col: 4, colOff: 0, colShare: 0, row: 4, rowOff: 0, rowShare: 0 },
  });
  assert.deepEqual([comments[1].row, comments[1].col, comments[1].author, comments[1].text, comments[1].visible], [4, 2, 'Bora', 'ikinci', true]);
  assert.deepEqual(comments[2], { row: 6, col: 0, text: '' });

  // The view finds them by cell, and draws the box of the one that is always shown.
  const layer = new SheetDrawings(sheet);
  assert.equal(layer.has(1, 1), true);
  assert.equal(layer.commentAt(4, 2)?.text, 'ikinci');
  assert.equal(layer.rects(createSheetMetrics(sheet)).length, 2);
});

test('xls comments: text in wide characters and text that goes on in a second CONTINUE record', () => {
  const wideText = 'Açıklama: şöyle';
  const wideBytes: number[] = [];
  for (let i = 0; i < wideText.length; i++) wideBytes.push(wideText.charCodeAt(i) & 0xff, wideText.charCodeAt(i) >> 8);
  const runs = [...u16(0), ...u16(0), ...u32(0), ...u16(wideText.length), ...u16(0), ...u32(0)];
  const text = [
    rec(0x01b6, u16(0x0212), u16(0), new Array(6).fill(0), u16(wideText.length), u16(runs.length), u32(0)),
    // The first ten characters two bytes each, the rest in a record of its own with its own width byte.
    rec(0x003c, [1], wideBytes.slice(0, 20)),
    rec(0x003c, [1], wideBytes.slice(20)),
    rec(0x003c, runs),
  ];
  const records = cat(drawingRecords([shape({ kind: 202, ot: 0x19, id: 1, text })]), note(0, 0, 1, 'Ayşe'));
  const comment = book(records).sheets[0].comments?.[0];
  assert.equal(comment?.text, wideText);
  assert.equal(comment?.author, 'Ayşe');
});

// --- damage ------------------------------------------------------------------------------

test('xls drawings: damaged drawing records never stop a sheet from opening', () => {
  const image = png(4, 4, RED);
  const good = drawingRecords([shape({ kind: 75, ot: 8, id: 1, props: [[0x4104, 1]] }), shape({ kind: 202, ot: 6, id: 2, text: txo('metin') })]);
  const tail = rec(0x00eb, drawingGroup([bse(image)]));
  let opened = 0;
  // Every single byte of the drawing records changed, one at a time and several ways.
  for (let i = 0; i < good.length; i++) {
    for (const value of [0x00, 0xff, good[i] ^ 0x80]) {
      const bad = good.slice();
      bad[i] = value;
      let workbook: Workbook | undefined;
      try {
        workbook = book(cat(number(0, 0, 1), bad, number(1, 0, 2)), tail);
      } catch (error) {
        // A record header that now claims more bytes than the stream has is an error of the record stream itself.
        assert.match(String(error), /Invalid xls|not supported/);
        continue;
      }
      opened++;
      const sheet = workbook.sheets[0];
      new SheetDrawings(sheet).rects(createSheetMetrics(sheet));
    }
  }
  assert.ok(opened > good.length * 2);
});

test('xls drawings: a truncated drawing, a store that is not one, and text that claims more than it has', async () => {
  const good = drawingRecords([shape({ kind: 75, ot: 8, id: 1, props: [[0x4104, 1]] })]);
  // The drawing container claims more bytes than there are.
  const short = rec(0x00ec, good.subarray(4, 30));
  const sheet = book(cat(short, obj(8, 1), number(2, 2, 9)), rec(0x00eb, [1, 2, 3, 4, 5, 6, 7, 8, 9])).sheets[0];
  assert.equal(sheet.drawings, undefined);
  assert.equal(getCell(sheet, 2, 2)?.value, 9);

  // A picture that names an entry of a store that could not be read.
  const noStore = book(good, rec(0x00eb, new Array(40).fill(0xff)));
  assert.equal((noStore.sheets[0].drawings?.[0].object as DrawingPicture).media, undefined);
  assert.equal(noStore.media, undefined);

  // A text that says it has 60,000 characters and has five.
  const liar = [rec(0x01b6, u16(0x0212), u16(0), new Array(6).fill(0), u16(60000), u16(16), u32(0)), rec(0x003c, [0], narrow('kisa!'))];
  const lied = book(cat(drawingRecords([shape({ kind: 202, ot: 0x19, id: 1, text: liar })]), note(0, 0, 1, 'a')));
  assert.equal(lied.sheets[0].comments?.[0].text, '');

  // Object records without a drawing, and a drawing without object records.
  const orphan = book(cat(obj(8, 1), ...txo('x'), note(0, 0, 1, 'a')));
  assert.deepEqual(orphan.sheets[0].comments, [{ row: 0, col: 0, text: '', author: 'a' }]);
  const bare = drawingRecords([shape({ kind: 1, ot: 2, id: 1 })]);
  const firstLength = bare[2] | (bare[3] << 8);
  assert.equal(book(bare.subarray(0, 4 + firstLength)).sheets[0].drawings?.length, 1);
});

test('xls drawings: groups nested without end are read as far as the limit goes', () => {
  let inner: Segment[][] = [shape({ kind: 1, ot: 2, id: 1, anchor: childAnchor(0, 0, 10, 10) })];
  for (let i = 0; i < 40; i++) inner = [groupOf(100 + i, [0, 0, 10, 10], i === 39 ? clientAnchor([0, 0, 0, 0], [2, 0, 2, 0]) : childAnchor(0, 0, 10, 10), inner)];
  const sheet = book(drawingRecords(inner)).sheets[0];
  let depth = 0;
  let object = sheet.drawings?.[0].object;
  while (object && object.type === 'group' && object.children.length > 0) {
    object = object.children[0];
    depth++;
  }
  assert.ok(depth > 0 && depth <= 17, `depth ${depth}`);
});

test('xls drawings: the older format keeps its objects another way; they are counted as not read', () => {
  const records = cat(rec(0x00ec, [1, 2, 3]), rec(0x005d, new Array(30).fill(0)), rec(0x001c, u16(0), u16(0), u16(2), narrow('ab')));
  const { workbook, report } = readXlsStream(workbookStream({ version: 0x0500, head: [], sheets: [{ name: 'S', records }] }), 'Book');
  assert.equal(workbook.sheets[0].drawings, undefined);
  assert.equal(workbook.sheets[0].comments, undefined);
  assert.equal(report.skippedRecords.get(0x00ec), 1);
  assert.equal(report.skippedRecords.get(0x005d), 1);
  assert.equal(report.comments, 1);
});

test('xls drawings: through open(), a file shows its picture', async () => {
  const image = png(4, 4, RED);
  const stream = workbookStream({
    head: cat(defaultStyles(), xf({ parent: 0 })),
    tail: rec(0x00eb, drawingGroup([bse(image)])),
    sheets: [{ name: 'S', records: drawingRecords([shape({ kind: 75, ot: 8, id: 1, props: [[0x4104, 1]] })]) }],
  });
  const file = await open(xlsFile(stream));
  const workbook = await file.workbook();
  assert.equal(workbook.sheets[0].drawings?.length, 1);
  assert.deepEqual(await workbook.media?.readPart('xls:blip:1'), image);
});
