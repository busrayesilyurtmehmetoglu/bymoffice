// Pictures added to a workbook, moved, resized and deleted (ops/pictures.ts) and
// what the writer makes of them: a new picture gets a media part, a drawing part
// when the sheet has none, the relationships and content types that lead to it;
// a picture of the file changes in its anchor only; whatever is undone leaves the
// file byte for byte as it was. Packages are read back by the engine and by
// test/tools/package_check.py.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { buildXlsx as buildWorkbook } from '../src/formats/xlsx/build';
import { encodeOp } from '../src/collab/wire';
import { CollabUnsupportedError } from '../src/collab/types';
import { createBlank, open, ZipArchive, type DrawingAnchor, type DrawingPicture, type Workbook } from '../src/index';
import { isAddedPicture, pictureSource } from '../src/model/sheet/pictures';
import { applyOp } from '../src/ops/ops';
import { addPicture, anchorLike, checkPicture, MAX_PICTURE_BYTES, movePicture, movePictureTo, picturesProblem, removePicture, sheetPictures } from '../src/ops/pictures';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import { apply, changedParts, hasPython, load, packageProblems, partText, save, sameBytes } from './edit-commands-fixtures';
import { bmpHeader, gifHeader, jpegHeader, png, pngHeader, pictureWorkbook } from './picture-fixtures';
import { pic } from './sheet-drawing-fixtures';
import { ratioRect } from '../src/render/sheet-drawings';
import { editDrawingPart } from '../src/formats/xlsx/chart-write';

const EMU = 9525;
const python = hasPython(['xml.etree.ElementTree', 'zipfile']);

async function blank() {
  return load(await createBlank('xlsx', { locale: 'tr' }));
}

const picturesOf = (wb: Workbook, sheet = 0) => (wb.sheets[sheet].drawings || []).filter((drawing) => drawing.object.type === 'picture');

// ------------------------------------------------------------------ the bytes

test('checkPicture: PNG, JPEG, GIF and BMP are read by their bytes, with their size', () => {
  assert.deepEqual(checkPicture(png(4, 3)), { type: 'image/png', width: 4, height: 3 });
  assert.deepEqual(checkPicture(jpegHeader(640, 480), 'image/jpg'), { type: 'image/jpeg', width: 640, height: 480 });
  assert.deepEqual(checkPicture(gifHeader(10, 20), 'IMAGE/GIF'), { type: 'image/gif', width: 10, height: 20 });
  assert.deepEqual(checkPicture(bmpHeader(7, 9)), { type: 'image/bmp', width: 7, height: 9 });
  // A BMP stored top down states a negative height.
  assert.deepEqual(checkPicture(bmpHeader(7, -9)), { type: 'image/bmp', width: 7, height: 9 });
});

test('checkPicture: a BMP with the OS/2 header of 12 bytes gives its size in two 16-bit words', () => {
  const core = new Uint8Array(26 + 4 * 3 * 3);
  const view = new DataView(core.buffer);
  core[0] = 0x42;
  core[1] = 0x4d;
  view.setUint32(2, core.length, true);
  view.setUint32(10, 26, true);
  view.setUint32(14, 12, true);
  view.setUint16(18, 4, true);
  view.setUint16(20, 3, true);
  view.setUint16(22, 1, true);
  view.setUint16(24, 24, true);
  assert.deepEqual(checkPicture(core), { type: 'image/bmp', width: 4, height: 3 });
  view.setUint16(18, 40000, true);
  assert.match(checkPicture(core) as string, /16384 pixels a side/);
});

test('checkPicture: what is not a picture that can be added is said, never guessed', () => {
  assert.match(checkPicture(new Uint8Array(0)) as string, /bytes/);
  assert.match(checkPicture('png' as unknown as Uint8Array) as string, /bytes/);
  assert.match(checkPicture(new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"/>')) as string, /not a PNG, JPEG, GIF or BMP/);
  assert.match(checkPicture(png(2, 2), 'image/svg+xml') as string, /PNG, JPEG, GIF or BMP/);
  assert.match(checkPicture(png(2, 2), 'image/jpeg') as string, /another kind/);
  // A PNG signature without its header, a GIF that is no version, a BMP whose header size is no header's.
  const broken = png(2, 2);
  broken[12] = 0x41;
  assert.match(checkPicture(broken) as string, /cannot be read/);
  assert.match(checkPicture(new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x30, 0x30, 1, 0, 1, 0, 0, 0, 0])) as string, /cannot be read/);
  const bmp = bmpHeader(2, 2);
  bmp[14] = 3;
  assert.match(checkPicture(bmp) as string, /cannot be read/);
  // A size of nothing, and a size beyond what is shown.
  assert.match(checkPicture(pngHeader(0, 10)) as string, /cannot be read/);
  assert.match(checkPicture(pngHeader(16385, 10)) as string, /16384 pixels a side/);
  assert.match(checkPicture(pngHeader(16000, 16000)) as string, /million pixels/);
  assert.match(checkPicture(jpegHeader(10, 20000)) as string, /16384/);
});

test('checkPicture: a file larger than 32 MB is refused before anything of it is read', () => {
  const big = new Uint8Array(MAX_PICTURE_BYTES + 1);
  big.set(png(2, 2));
  assert.match(checkPicture(big) as string, /32 MB/);
});

test('addPicture: a refused picture changes nothing', async () => {
  const doc = await blank();
  assert.throws(() => addPicture(doc.wb, 0, new TextEncoder().encode('nothing of a picture')), RangeError);
  assert.throws(() => addPicture(doc.wb, 0, png(2, 2), 'image/webp'), RangeError);
  assert.throws(() => addPicture(doc.wb, 0, png(2, 2), undefined, undefined, { description: 'a\u0001b' }), /character/);
  assert.throws(() => addPicture(doc.wb, 3, png(2, 2)), /No sheet/);
  assert.equal(doc.wb.sheets[0].drawings, undefined);
});

// ------------------------------------------------------------------ added pictures

test('a picture added to a blank workbook: a media part, a drawing part, their relationships and content types', async () => {
  const doc = await blank();
  const image = png(4, 3);
  apply(doc, addPicture(doc.wb, 0, image, 'image/png', { top: 1, left: 1, bottom: 4, right: 3 }, { description: 'Kaşe & imza' }));
  const [drawing] = picturesOf(doc.wb);
  assert.ok(isAddedPicture(drawing.object));
  // The bytes were copied: what the caller does with its array afterwards changes nothing.
  image[20] ^= 0xff;
  assert.notEqual((drawing.object as DrawingPicture).data!.bytes[20], image[20]);
  const saved = await save(doc);
  const zip = ZipArchive.read(saved);
  assert.ok(zip.has('xl/media/image1.png'));
  assert.ok(sameBytes(await zip.read('xl/media/image1.png'), (drawing.object as DrawingPicture).data!.bytes));
  const drawingXml = await partText(saved, 'xl/drawings/drawing1.xml');
  assert.match(drawingXml, /<xdr:twoCellAnchor><xdr:from><xdr:col>1<\/xdr:col><xdr:colOff>0<\/xdr:colOff><xdr:row>1<\/xdr:row>/);
  assert.match(drawingXml, /<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Picture 1" descr="Kaşe &amp; imza"\/>/);
  assert.match(drawingXml, /<a:blip xmlns:r="http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships" r:embed="rId1"\/>/);
  assert.match(await partText(saved, 'xl/drawings/_rels/drawing1.xml.rels'), /Id="rId1" Type="http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships\/image" Target="..\/media\/image1.png"/);
  assert.match(await partText(saved, 'xl/worksheets/_rels/sheet1.xml.rels'), /relationships\/drawing" Target="..\/drawings\/drawing1.xml"/);
  assert.match(await partText(saved, 'xl/worksheets/sheet1.xml'), /<drawing r:id="rId\d+"\/>/);
  const types = await partText(saved, '[Content_Types].xml');
  assert.match(types, /PartName="\/xl\/media\/image1.png" ContentType="image\/png"|Extension="png" ContentType="image\/png"/);
  assert.match(types, /PartName="\/xl\/drawings\/drawing1.xml" ContentType="application\/vnd.openxmlformats-officedocument.drawing\+xml"/);
  // Read back: a picture of the file, at the same place, with its bytes.
  const again = await load(saved);
  const [read] = picturesOf(again.wb);
  const object = read.object as DrawingPicture;
  assert.equal(object.description, 'Kaşe & imza');
  assert.deepEqual(read.anchor, { kind: 'twoCell', from: { col: 1, colOff: 0, row: 1, rowOff: 0 }, to: { col: 4, colOff: 0, row: 5, rowOff: 0 } });
  assert.ok(sameBytes(await again.wb.media!.readPart(object.media!), (drawing.object as DrawingPicture).data!.bytes));
  if (python) assert.deepEqual(packageProblems({ 'added.xlsx': saved }), []);
});

test('a picture added without a place stands at A1 in its own size, moves with its cell and keeps its size', async () => {
  const doc = await blank();
  apply(doc, addPicture(doc.wb, 0, png(100, 50)));
  const [drawing] = picturesOf(doc.wb);
  assert.equal(drawing.anchor.kind, 'twoCell');
  assert.equal(drawing.anchor.editAs, 'oneCell');
  const rect = anchorLike(drawing.anchor, createSheetMetrics(doc.wb.sheets[0]), { x: 0, y: 0, width: 100, height: 50 });
  assert.deepEqual(drawing.anchor, rect);
  const xml = await partText(await save(doc), 'xl/drawings/drawing1.xml');
  assert.match(xml, /<xdr:twoCellAnchor editAs="oneCell">/);
  assert.match(xml, /<a:ext cx="952500" cy="476250"\/>/);
  // A large picture is made smaller, its sides in proportion.
  apply(doc, addPicture(doc.wb, 0, pngHeader(4000, 1000)));
  const big = picturesOf(doc.wb)[1];
  const box = createSheetMetrics(doc.wb.sheets[0]);
  const from = big.anchor.from!;
  const to = big.anchor.to!;
  const width = box.cols.offset(to.col) + to.colOff / EMU - (box.cols.offset(from.col) + from.colOff / EMU);
  assert.ok(Math.abs(width - 800) <= 1, `width ${width}`);
});

test('the same bytes added twice are one media part; a picture of other bytes is another', async () => {
  const doc = await blank();
  const image = png(4, 3);
  apply(doc, addPicture(doc.wb, 0, image, undefined, { top: 0, left: 0, bottom: 2, right: 2 }));
  apply(doc, addPicture(doc.wb, 0, image.slice(), undefined, { top: 5, left: 0, bottom: 7, right: 2 }));
  apply(doc, addPicture(doc.wb, 0, png(4, 3, [0, 0, 255]), undefined, { top: 9, left: 0, bottom: 11, right: 2 }));
  const saved = await save(doc);
  const zip = ZipArchive.read(saved);
  assert.deepEqual(
    zip.names().filter((name) => name.startsWith('xl/media/')),
    ['xl/media/image1.png', 'xl/media/image2.png'],
  );
  const xml = await partText(saved, 'xl/drawings/drawing1.xml');
  assert.deepEqual([...xml.matchAll(/r:embed="(rId\d+)"/g)].map((m) => m[1]), ['rId1', 'rId1', 'rId2']);
  assert.deepEqual([...xml.matchAll(/cNvPr id="(\d+)"/g)].map((m) => m[1]), ['2', '3', '4']);
  const rels = await partText(saved, 'xl/drawings/_rels/drawing1.xml.rels');
  assert.equal([...rels.matchAll(/<Relationship /g)].length, 2);
  const again = await load(saved);
  assert.equal(picturesOf(again.wb).length, 3);
  if (python) assert.deepEqual(packageProblems({ 'twice.xlsx': saved }), []);
});

test('added, then undone: the file is byte for byte as it was; saved in between, the next save takes the parts out again', async () => {
  const file = await pictureWorkbook();
  const doc = await load(file);
  apply(doc, addPicture(doc.wb, 0, png(3, 3)));
  doc.history.undo();
  assert.ok(sameBytes(await save(doc), file), 'undone before a save');
  doc.history.redo();
  const once = await save(doc);
  assert.ok(ZipArchive.read(once).has('xl/media/image2.png'));
  doc.history.undo();
  assert.ok(sameBytes(await save(doc), file), 'undone after a save');
  doc.history.redo();
  const twice = await save(doc);
  assert.ok(sameBytes(twice, once), 'done again: the parts of the first save, under the same names');
});

test('a picture added to a sheet with a drawing goes into that drawing, above the others, with an id no object has', async () => {
  const file = await pictureWorkbook();
  const doc = await load(file);
  apply(doc, addPicture(doc.wb, 0, png(3, 3), undefined, { top: 10, left: 0, bottom: 12, right: 1 }, { name: 'Yeni' }));
  const saved = await save(doc);
  // The file declares PNG pictures by their extension: the content types stay as they are.
  assert.deepEqual(await changedParts(file, saved), ['xl/drawings/_rels/drawing1.xml.rels', 'xl/drawings/drawing1.xml', 'xl/media/image2.png']);
  const before = await partText(file, 'xl/drawings/drawing1.xml');
  const after = await partText(saved, 'xl/drawings/drawing1.xml');
  // What was there is there with its very bytes, and the new picture comes last.
  assert.ok(after.startsWith(before.slice(0, before.indexOf('</xdr:wsDr>'))));
  assert.match(after.slice(before.indexOf('</xdr:wsDr>')), /^<xdr:twoCellAnchor><xdr:from>.*<xdr:cNvPr id="5" name="Yeni"\/>.*r:embed="rId1".*<\/xdr:twoCellAnchor><\/xdr:wsDr>$/);
  assert.match(await partText(saved, 'xl/drawings/_rels/drawing1.xml.rels'), /Id="rId1" Type="[^"]+\/image" Target="..\/media\/image2.png"/);
  if (python) assert.deepEqual(packageProblems({ 'into.xlsx': saved }), []);
});

test('pictures of each kind are added from an anchor of that kind', async () => {
  const doc = await blank();
  const one: DrawingAnchor = { kind: 'oneCell', from: { col: 2, colOff: 9525, row: 3, rowOff: 0 }, width: 50 * EMU, height: 20 * EMU };
  const abs: DrawingAnchor = { kind: 'absolute', x: 10 * EMU, y: 20 * EMU, width: 30 * EMU, height: 40 * EMU };
  const two: DrawingAnchor = { kind: 'twoCell', editAs: 'absolute', from: { col: 0, colOff: 0, row: 0, rowOff: 0 }, to: { col: 1, colOff: 0, row: 1, rowOff: 0 } };
  for (const anchor of [one, abs, two]) apply(doc, addPicture(doc.wb, 0, gifHeader(5, 5), 'image/gif', anchor));
  const saved = await save(doc);
  const xml = await partText(saved, 'xl/drawings/drawing1.xml');
  assert.match(xml, /<xdr:oneCellAnchor><xdr:from><xdr:col>2<\/xdr:col><xdr:colOff>9525<\/xdr:colOff><xdr:row>3<\/xdr:row><xdr:rowOff>0<\/xdr:rowOff><\/xdr:from><xdr:ext cx="476250" cy="190500"\/><xdr:pic>/);
  assert.match(xml, /<xdr:absoluteAnchor><xdr:pos x="95250" y="190500"\/><xdr:ext cx="285750" cy="381000"\/><xdr:pic>/);
  assert.match(xml, /<xdr:twoCellAnchor editAs="absolute">/);
  assert.ok(ZipArchive.read(saved).has('xl/media/image1.gif'));
  const again = await load(saved);
  assert.deepEqual(picturesOf(again.wb).map((d) => d.anchor), [one, abs, two]);
  if (python) assert.deepEqual(packageProblems({ 'kinds.xlsx': saved }), []);
});

// ------------------------------------------------------------------ pictures of the file

test('pictures of the file are moved: their anchors change, nothing else of the drawing does, and each keeps its kind', async () => {
  const file = await pictureWorkbook({ editAs: 'oneCell' });
  const doc = await load(file);
  const metrics = createSheetMetrics(doc.wb.sheets[0]);
  const infos = sheetPictures(doc.wb, 0);
  assert.deepEqual(infos.map((info) => [info.anchor.kind, info.ofFile, info.removable]), [
    ['twoCell', true, true],
    ['oneCell', true, true],
    ['absolute', true, true],
  ]);
  for (const info of infos) apply(doc, movePictureTo(doc.wb, 0, info.id, { x: 640, y: 200 + info.id * 100, width: 70, height: 35 }, metrics));
  assert.deepEqual(
    sheetPictures(doc.wb, 0).map((info) => [info.anchor.kind, info.anchor.editAs]),
    [
      ['twoCell', 'oneCell'],
      ['oneCell', undefined],
      ['absolute', undefined],
    ],
  );
  const saved = await save(doc);
  assert.deepEqual(await changedParts(file, saved), ['xl/drawings/drawing1.xml']);
  const xml = await partText(saved, 'xl/drawings/drawing1.xml');
  // What the engine has no field for is still there, and so is every name.
  assert.match(xml, /<a14:useLocalDpi xmlns:a14="http:\/\/schemas.microsoft.com\/office\/drawing\/2010\/main" val="0"\/>/);
  assert.match(xml, /<xdr:twoCellAnchor editAs="oneCell">/);
  assert.match(xml, /<xdr:absoluteAnchor><xdr:pos x="6096000" y="3810000"\/><xdr:ext cx="666750" cy="333375"\/>/);
  const again = await load(saved);
  const now = sheetPictures(again.wb, 0);
  assert.deepEqual(now.map((info) => info.name), ['Logo', 'İkinci', 'Üçüncü']);
  for (const info of now) {
    const rect = anchorLike(info.anchor, metrics, { x: 640, y: 200 + info.id * 100, width: 70, height: 35 });
    assert.deepEqual(info.anchor, rect);
  }
  if (python) assert.deepEqual(packageProblems({ 'moved.xlsx': saved }), []);
});

test('movePicture: the cells it covers, or an anchor of its own kind; another kind is refused; the same place is no step', async () => {
  const doc = await load(await pictureWorkbook());
  const two = movePicture(doc.wb, 0, 0, { top: 3, left: 3, bottom: 5, right: 5 });
  assert.equal(two.length, 1);
  apply(doc, movePicture(doc.wb, 0, 1, { top: 3, left: 3, bottom: 5, right: 5 }));
  assert.equal(sheetPictures(doc.wb, 0)[1].anchor.kind, 'oneCell');
  assert.throws(() => movePicture(doc.wb, 0, 2, { kind: 'oneCell', from: { col: 0, colOff: 0, row: 0, rowOff: 0 }, width: 1, height: 1 }), /keeps the kind/);
  assert.deepEqual(movePicture(doc.wb, 0, 2, sheetPictures(doc.wb, 0)[2].anchor), []);
  assert.throws(() => movePicture(doc.wb, 0, 7, { top: 0, left: 0, bottom: 1, right: 1 }), /No picture at index 7/);
});

test('a picture of the file is deleted: the relationship and the media part go once nothing else shows them; undone, the file is as it was', async () => {
  const file = await pictureWorkbook({ shared: true });
  const doc = await load(file);
  // Two of the three show image1.png through rIdI1, one through rIdI2.
  apply(doc, removePicture(doc.wb, 0, 1));
  let saved = await save(doc);
  assert.ok(ZipArchive.read(saved).has('xl/media/image1.png'), 'the others still show it');
  assert.doesNotMatch(await partText(saved, 'xl/drawings/_rels/drawing1.xml.rels'), /rIdI2/);
  apply(doc, removePicture(doc.wb, 0, 0));
  apply(doc, removePicture(doc.wb, 0, 0));
  assert.equal(doc.wb.sheets[0].drawings, undefined);
  saved = await save(doc);
  const zip = ZipArchive.read(saved);
  assert.ok(!zip.has('xl/media/image1.png'), 'nothing shows it any more');
  assert.ok(!zip.has('xl/drawings/drawing1.xml'), 'nothing is left to draw');
  assert.doesNotMatch(await partText(saved, 'xl/worksheets/sheet1.xml'), /<drawing /);
  if (python) assert.deepEqual(packageProblems({ 'removed.xlsx': saved }), []);
  doc.history.undo();
  doc.history.undo();
  doc.history.undo();
  assert.ok(sameBytes(await save(doc), file), 'every deletion undone: byte for byte');
});

test('a picture that shares its wrapper of alternatives with another drawing is moved, not deleted', async () => {
  const both = pic('rIdI1', { name: 'A' });
  const wrapped =
    '<mc:AlternateContent><mc:Choice Requires="a14">' +
    `${'<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'}${both}<xdr:clientData/></xdr:twoCellAnchor>` +
    `${'<xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'}${both.replace('cNvPr id="2"', 'cNvPr id="3"')}<xdr:clientData/></xdr:twoCellAnchor>` +
    '</mc:Choice><mc:Fallback/></mc:AlternateContent>';
  const doc = await load(await pictureWorkbook({ drawing: wrapped }));
  assert.deepEqual(sheetPictures(doc.wb, 0).map((info) => info.removable), [false, false]);
  assert.throws(() => removePicture(doc.wb, 0, 0), /shares its place/);
  apply(doc, movePicture(doc.wb, 0, 1, { top: 4, left: 4, bottom: 6, right: 6 }));
  const saved = await save(doc);
  assert.match(await partText(saved, 'xl/drawings/drawing1.xml'), /<mc:AlternateContent><mc:Choice Requires="a14">.*<xdr:col>4<\/xdr:col>.*<\/mc:Choice><mc:Fallback\/><\/mc:AlternateContent>/);
});

test('a protected sheet whose objects are protected takes no picture; one that only protects its cells does', async () => {
  const locked = await load(await pictureWorkbook({ protection: '<sheetProtection sheet="1" objects="1" scenarios="1"/>' }));
  assert.match(picturesProblem(locked.wb, 0)!, /protected/);
  assert.throws(() => addPicture(locked.wb, 0, png(2, 2)), /protected/);
  assert.throws(() => movePicture(locked.wb, 0, 0, { top: 4, left: 4, bottom: 6, right: 6 }), /protected/);
  assert.throws(() => removePicture(locked.wb, 0, 0), /protected/);
  const cells = await load(await pictureWorkbook({ protection: '<sheetProtection sheet="1"/>' }));
  assert.equal(picturesProblem(cells.wb, 0), undefined);
  apply(cells, removePicture(cells.wb, 0, 2));
});

test('a sheet whose drawing takes no change takes no picture, and its pictures stay where they are', async () => {
  const doc = await load(await pictureWorkbook());
  doc.wb.sheets[0].drawingsFixed = 'its drawing is shared with another sheet';
  assert.throws(() => addPicture(doc.wb, 0, png(2, 2)), /shared with another sheet/);
  assert.throws(() => removePicture(doc.wb, 0, 0), /shared with another sheet/);
});

// ------------------------------------------------------------------ rows, columns and sheets

test('rows and columns inserted and deleted move added pictures by the kind of their anchor, as they move those of the file', async () => {
  const doc = await blank();
  const at = (row: number): DrawingAnchor => ({ kind: 'twoCell', from: { col: 1, colOff: 0, row, rowOff: 0 }, to: { col: 3, colOff: 0, row: row + 2, rowOff: 0 } });
  apply(doc, addPicture(doc.wb, 0, png(2, 2), undefined, at(5)));
  apply(doc, addPicture(doc.wb, 0, png(2, 2), undefined, { ...at(5), editAs: 'oneCell' }));
  apply(doc, addPicture(doc.wb, 0, png(2, 2), undefined, { ...at(5), editAs: 'absolute' }));
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 3 }]);
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 9, count: 1 }]);
  const now = sheetPictures(doc.wb, 0).map((info) => [info.anchor.from!.row, info.anchor.to!.row]);
  assert.deepEqual(now, [
    [8, 11],
    [8, 10],
    [5, 7],
  ]);
  const saved = await save(doc);
  const again = await load(saved);
  assert.deepEqual(
    sheetPictures(again.wb, 0).map((info) => [info.anchor.from!.row, info.anchor.to!.row]),
    now,
  );
  // A picture moved after the rows were inserted stands where it was put; only later edits move it.
  apply(doc, movePicture(doc.wb, 0, 0, at(20)));
  apply(doc, [{ type: 'delete', sheet: 0, axis: 'row', at: 0, count: 2 }]);
  assert.deepEqual([sheetPictures(doc.wb, 0)[0].anchor.from!.row, sheetPictures(doc.wb, 0)[0].anchor.to!.row], [18, 20]);
  if (python) assert.deepEqual(packageProblems({ 'rows.xlsx': await save(doc) }), []);
});

test('a picture of the file moves with rows deleted under it, and a sheet with an added picture can be renamed and deleted', async () => {
  const file = await pictureWorkbook();
  const doc = await load(file);
  apply(doc, addPicture(doc.wb, 0, png(2, 2)));
  apply(doc, [{ type: 'delete', sheet: 0, axis: 'col', at: 0, count: 1 }]);
  assert.deepEqual(sheetPictures(doc.wb, 0)[0].anchor.from, { col: 0, colOff: 0, row: 1, rowOff: 0 });
  apply(doc, [{ type: 'insertSheet', sheet: 1, name: 'İkinci' }]);
  apply(doc, [{ type: 'renameSheet', sheet: 0, name: 'Resimli' }]);
  let saved = await save(doc);
  if (python) assert.deepEqual(packageProblems({ 'renamed.xlsx': saved }), []);
  apply(doc, [{ type: 'deleteSheet', sheet: 0 }]);
  saved = await save(doc);
  const zip = ZipArchive.read(saved);
  assert.ok(!zip.names().some((name) => name.startsWith('xl/media/')), 'the pictures went with their sheet');
  if (python) assert.deepEqual(packageProblems({ 'deleted.xlsx': saved }), []);
  for (let i = 0; i < 5; i++) doc.history.undo();
  assert.ok(sameBytes(await save(doc), file), 'all undone: byte for byte');
});

test('a picture added to a sheet that is in no file is written with it', async () => {
  const doc = await load(await pictureWorkbook());
  apply(doc, [{ type: 'insertSheet', sheet: 1, name: 'Yeni' }]);
  apply(doc, addPicture(doc.wb, 1, jpegHeader(30, 20), 'image/jpeg', { top: 0, left: 0, bottom: 3, right: 3 }));
  const saved = await save(doc);
  const again = await load(saved);
  const [picture] = picturesOf(again.wb, 1);
  assert.ok(picture);
  const media = (picture.object as DrawingPicture).media!;
  assert.match(media, /\.jpeg$/);
  assert.ok(sameBytes(await again.wb.media!.readPart(media), jpegHeader(30, 20)));
  if (python) assert.deepEqual(packageProblems({ 'new-sheet.xlsx': saved }), []);
});

// ------------------------------------------------------------------ elsewhere

test('a workbook built anew (saveAsXlsx) has the pictures that were added, with their bytes', async () => {
  const doc = await blank();
  apply(doc, addPicture(doc.wb, 0, png(5, 5), undefined, { top: 0, left: 0, bottom: 2, right: 2 }));
  const built = await buildWorkbook(doc.wb);
  const zip = ZipArchive.read(built);
  const media = zip.names().filter((name) => name.startsWith('xl/media/'));
  assert.equal(media.length, 1);
  assert.ok(sameBytes(await zip.read(media[0]), (picturesOf(doc.wb)[0].object as DrawingPicture).data!.bytes));
  if (python) assert.deepEqual(packageProblems({ 'built.xlsx': built }), []);
});

test('pictureSource: added pictures from the model, those of the file from the file, a picture taken away from nowhere', async () => {
  const doc = await load(await pictureWorkbook());
  apply(doc, addPicture(doc.wb, 0, png(2, 2)));
  const source = pictureSource(doc.wb);
  const added = picturesOf(doc.wb)[3].object as DrawingPicture;
  assert.ok(sameBytes(await source.readPart(added.media!), added.data!.bytes));
  const ofFile = picturesOf(doc.wb)[0].object as DrawingPicture;
  assert.ok(sameBytes(await source.readPart(ofFile.media!), png(4, 3)));
  doc.history.undo();
  await assert.rejects(source.readPart(added.media!));
  await assert.rejects(source.readPart('/xl/worksheets/sheet1.xml'));
});

test('a step that takes an added picture away keeps its bytes in its inverse, and says how much they weigh', async () => {
  const doc = await blank();
  apply(doc, addPicture(doc.wb, 0, png(300, 300)));
  const bytes = (picturesOf(doc.wb)[0].object as DrawingPicture).data!.bytes.length;
  const inverse = applyOp(doc.wb, removePicture(doc.wb, 0, 0)[0]);
  assert.equal(inverse.type, 'setDrawings');
  assert.equal((inverse as { weight?: number }).weight, Math.ceil(bytes / 1024));
  // Moving one keeps nothing alive that the model does not hold.
  applyOp(doc.wb, inverse);
  const moved = applyOp(doc.wb, movePicture(doc.wb, 0, 0, { top: 3, left: 3, bottom: 5, right: 5 })[0]);
  assert.equal((moved as { weight?: number }).weight, undefined);
});

test('in a room, pictures are refused before anything is changed, with a code of their own', async () => {
  const doc = await load(await pictureWorkbook());
  for (const ops of [addPicture(doc.wb, 0, png(2, 2)), movePicture(doc.wb, 0, 0, { top: 4, left: 4, bottom: 6, right: 6 }), removePicture(doc.wb, 0, 0)]) {
    assert.throws(
      () => encodeOp(doc.wb, ops[0]),
      (error: unknown) => error instanceof CollabUnsupportedError && error.what === 'picture',
    );
  }
});

test('ratioRect: a corner keeps the proportions of the box, the corner across from it stays; a side or a move is left as it is', () => {
  const start = { x: 100, y: 100, width: 200, height: 100 };
  assert.deepEqual(ratioRect(start, { x: 100, y: 100, width: 300, height: 110 }, 'se', 4), { x: 100, y: 100, width: 300, height: 150 });
  assert.deepEqual(ratioRect(start, { x: 100, y: 50, width: 210, height: 150 }, 'ne', 4), { x: 100, y: 50, width: 300, height: 150 });
  assert.deepEqual(ratioRect(start, { x: 0, y: 100, width: 300, height: 100 }, 'sw', 4), { x: 0, y: 100, width: 300, height: 150 });
  assert.deepEqual(ratioRect(start, { x: 290, y: 190, width: 10, height: 10 }, 'nw', 4), { x: 280, y: 190, width: 20, height: 10 });
  // Not past the corner of the sheet.
  assert.deepEqual(ratioRect(start, { x: 0, y: 0, width: 300, height: 200 }, 'nw', 4), { x: 0, y: 50, width: 300, height: 150 });
  const side = { x: 100, y: 100, width: 250, height: 100 };
  assert.equal(ratioRect(start, side, 'e', 4), side);
  assert.equal(ratioRect(start, side, 'move', 4), side);
});

test('a picture added to a drawing part of other prefixes is written in the prefixes of the part, its texts escaped', () => {
  const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const source = `<?xml version="1.0" encoding="UTF-8"?>\n<wsDr xmlns="${NS_XDR}" xmlns:d="${NS_A}"><absoluteAnchor><pos x="0" y="0"/><ext cx="9525" cy="9525"/><sp><nvSpPr><cNvPr id="7" name="S"/><cNvSpPr/></nvSpPr><spPr/></sp><clientData/></absoluteAnchor></wsDr>`;
  const anchor: DrawingAnchor = { kind: 'oneCell', from: { col: 1, colOff: 0, row: 2, rowOff: 0 }, width: 95250, height: 95250 };
  const added = [{ anchor, relId: 'rId9', name: 'P & Q', description: '"alt" <b>', picture: { box: { x: 1, y: 2, width: 3, height: 4 } } }];
  const out = new TextDecoder().decode(editDrawingPart(new TextEncoder().encode(source), { moved: new Map(), removed: new Set(), added }).bytes);
  // What was there keeps its bytes; the new anchor uses the default namespace of the part and its "d" prefix for drawings.
  assert.ok(out.startsWith(source.slice(0, source.indexOf('</wsDr>'))));
  assert.match(out, /<oneCellAnchor><from><col>1<\/col><colOff>0<\/colOff><row>2<\/row><rowOff>0<\/rowOff><\/from><ext cx="95250" cy="95250"\/><pic><nvPicPr><cNvPr id="8" name="P &amp; Q" descr="&quot;alt&quot; &lt;b>"\/><cNvPicPr><d:picLocks noChangeAspect="1"\/><\/cNvPicPr><\/nvPicPr>/);
  assert.match(out, /<d:blip xmlns:r="http:\/\/schemas.openxmlformats.org\/officeDocument\/2006\/relationships" r:embed="rId9"\/>/);
  assert.match(out, /<d:off x="1" y="2"\/><d:ext cx="3" cy="4"\/>/);
});

test('a picture of the file whose part a save took out is read again once its deletion is undone: by the view, by a workbook built anew, by the next save', async () => {
  const file = await pictureWorkbook();
  const doc = await load(file);
  for (let i = 0; i < 3; i++) apply(doc, removePicture(doc.wb, 0, 0));
  assert.ok(!ZipArchive.read(await save(doc)).has('xl/media/image1.png'), 'the save took the part out');
  for (let i = 0; i < 3; i++) doc.history.undo();
  const key = (doc.wb.sheets[0].drawings![0].object as DrawingPicture).media!;
  assert.ok(sameBytes(await pictureSource(doc.wb).readPart(key), png(4, 3)));
  assert.ok(sameBytes(await doc.wb.media!.readPart(key), png(4, 3)));
  const built = ZipArchive.read(await buildWorkbook(doc.wb));
  const media = built.names().filter((name) => name.startsWith('xl/media/'));
  assert.equal(media.length, 1);
  assert.ok(sameBytes(await built.read(media[0]), png(4, 3)));
  assert.ok(sameBytes(await save(doc), file), 'saved again: the file that was opened');
  assert.ok(sameBytes(await pictureSource(doc.wb).readPart(key), png(4, 3)), 'and the part is read from the package again');
});

test('the pictures of a sheet that was deleted, saved without it and brought back are read again', async () => {
  const doc = await load(await pictureWorkbook());
  apply(doc, [{ type: 'insertSheet', sheet: 1, name: 'B' }]);
  apply(doc, [{ type: 'deleteSheet', sheet: 0 }]);
  assert.ok(!ZipArchive.read(await save(doc)).has('xl/media/image1.png'));
  doc.history.undo();
  const key = (doc.wb.sheets[0].drawings![0].object as DrawingPicture).media!;
  assert.ok(sameBytes(await pictureSource(doc.wb).readPart(key), png(4, 3)));
});
