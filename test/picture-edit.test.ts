// Pictures of a sheet turned, cropped, described anew, copied, cut and pasted (ops/pictures.ts),
// the geometry the view sizes, turns and crops them with (render/picture-frame.ts), and what the
// writer makes of it: of a picture of the file only `a:xfrm rot`, `a:srcRect` and `descr` change
// besides its anchor; what is undone leaves the file byte for byte as it was. Saved packages are
// read back by the engine, by test/tools/package_check.py and by openpyxl with ElementTree.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { encodeOp } from '../src/collab/wire';
import { CollabUnsupportedError } from '../src/collab/types';
import { createBlank, ZipArchive, type DrawingPicture, type Workbook } from '../src/index';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import { editDrawingPart } from '../src/formats/xlsx/chart-write';
import {
  addPicture,
  checkedCrop,
  copyPicture,
  cropPicture,
  cutPicture,
  describePicture,
  pastePicture,
  pictureCopyProblem,
  pictureRotation,
  PICTURE_COPY_OFFSET,
  rotatePicture,
  sheetPictures,
  type PictureCopy,
} from '../src/ops/pictures';
import { anchorRect, drawingAnchorNow } from '../src/render/sheet-drawings';
import { croppedBox, draggedCrop, localPoint, rotateHandlePoint, rotationAt, sizedFrame, turnedBox, uncroppedBox } from '../src/render/picture-frame';
import { apply, changedParts, hasPython, load, packageProblems, partText, runPython, save, sameBytes } from './edit-commands-fixtures';
import { png, pictureWorkbook } from './picture-fixtures';
import { pic } from './sheet-drawing-fixtures';

const python = hasPython(['xml.etree.ElementTree', 'zipfile']);
const openpyxl = hasPython(['openpyxl']);

async function blank() {
  return load(await createBlank('xlsx', { locale: 'tr' }));
}

const pictureOf = (wb: Workbook, id: number, sheet = 0) => wb.sheets[sheet].drawings![id].object as DrawingPicture;
const boxOf = (wb: Workbook, id: number, sheet = 0) => {
  const s = wb.sheets[sheet];
  return anchorRect(drawingAnchorNow(s.drawings![id], s.structuralEdits), createSheetMetrics(s));
};
const near = (a: number, b: number, what: string, within = 0.01) => assert.ok(Math.abs(a - b) <= within, `${what}: ${a} is not ${b}`);

/**
 * What a reader that is not this engine finds in the saved files: for each file the number of
 * pictures openpyxl reads on each sheet (when it is installed), and the pictures of every drawing
 * part as ElementTree reads them (name, alternative text, turn, crop, flips).
 */
function readOutside(files: { [name: string]: Uint8Array }): { [name: string]: { images: number[] | null; pictures: { name: string; descr: string | null; rot: string | null; crop: { [side: string]: string } | null; flipH: string | null }[] } } {
  return runPython(
    files,
    String.raw`
import json, sys, os, zipfile, warnings
import xml.etree.ElementTree as ET
warnings.simplefilter('ignore')
try:
    import openpyxl
except Exception:
    openpyxl = None
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
XDR = '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}'
out = {}
folder = sys.argv[1]
for name in sorted(os.listdir(folder)):
    path = os.path.join(folder, name)
    z = zipfile.ZipFile(path)
    pictures = []
    for part in sorted(n for n in z.namelist() if n.startswith('xl/drawings/') and n.endswith('.xml')):
        root = ET.fromstring(z.read(part))
        for el in root.iter(XDR + 'pic'):
            props = el.find(XDR + 'nvPicPr/' + XDR + 'cNvPr')
            xfrm = el.find(XDR + 'spPr/' + A + 'xfrm')
            rect = el.find(XDR + 'blipFill/' + A + 'srcRect')
            pictures.append({
                'name': props.get('name'),
                'descr': props.get('descr'),
                'rot': xfrm.get('rot') if xfrm is not None else None,
                'flipH': xfrm.get('flipH') if xfrm is not None else None,
                'crop': dict(rect.attrib) if rect is not None else None,
            })
    images = None
    if openpyxl is not None:
        book = openpyxl.load_workbook(path)
        images = [len(ws._images) for ws in book.worksheets]
    out[name] = {'images': images, 'pictures': pictures}
print(json.dumps(out))
`,
  );
}

// ------------------------------------------------------------------ geometry

test('turnedBox: near a quarter turn a box is turned about its middle; otherwise it is the box itself', () => {
  const box = { x: 100, y: 50, width: 200, height: 100 };
  assert.deepEqual(turnedBox(box, undefined), box);
  assert.deepEqual(turnedBox(box, 30), box);
  assert.deepEqual(turnedBox(box, 180), box);
  assert.deepEqual(turnedBox(box, 90), { x: 150, y: 0, width: 100, height: 200 });
  assert.deepEqual(turnedBox(box, 270), { x: 150, y: 0, width: 100, height: 200 });
  assert.deepEqual(turnedBox(turnedBox(box, 100), 100), box, 'turned twice: the box again');
  assert.deepEqual(turnedBox(box, 44.9), box);
  assert.deepEqual(turnedBox(box, 45), { x: 150, y: 0, width: 100, height: 200 });
});

test('sizedFrame: a box that is not turned is sized as any drawing; a turned one in its own frame, the side across staying put', () => {
  const box = { x: 100, y: 100, width: 200, height: 100 };
  assert.deepEqual(sizedFrame(box, 0, 'se', 20, 10, 4), { x: 100, y: 100, width: 220, height: 110 });
  assert.deepEqual(sizedFrame(box, 0, 'move', -500, 10, 4), { x: 0, y: 110, width: 200, height: 100 });
  // Turned a quarter: dragging the handle "e" (which is at the bottom on the screen) down makes the box wider.
  const turned = sizedFrame(box, 90, 'e', 0, 40, 4);
  near(turned.width, 240, 'width');
  near(turned.height, 100, 'height');
  // The side across (the top on the screen, its "w" side) stays where it was.
  const before = localPoint(box, 90, 0, 0);
  void before;
  const westBefore = { x: box.x + box.width / 2, y: box.y + box.height / 2 - box.width / 2 };
  const westAfter = { x: turned.x + turned.width / 2, y: turned.y + turned.height / 2 - turned.width / 2 };
  near(westAfter.x, westBefore.x, 'west x');
  near(westAfter.y, westBefore.y, 'west y');
  // With the proportions kept by a corner.
  const kept = sizedFrame(box, 30, 'se', 50, 0, 4, true);
  near(kept.width / kept.height, 2, 'ratio');
  // Never smaller than the least size.
  const small = sizedFrame(box, 60, 'nw', 1000, 1000, 4);
  assert.ok(small.width >= 4 - 1e-9 && small.height >= 4 - 1e-9);
});

test('rotationAt and rotateHandlePoint: straight above the middle is no turn, clockwise up to 359; Shift gives steps of 15', () => {
  const box = { x: 0, y: 0, width: 100, height: 100 };
  assert.equal(rotationAt(box, 50, -100, false), 0);
  assert.equal(rotationAt(box, 150, 50, false), 90);
  assert.equal(rotationAt(box, 50, 200, false), 180);
  assert.equal(rotationAt(box, -50, 50, false), 270);
  assert.equal(rotationAt(box, 100, 0, false), 45);
  assert.equal(rotationAt(box, 100, 3, true), 45);
  assert.equal(rotationAt(box, 51, -100, true), 0);
  assert.equal(rotationAt(box, 50, 50, false), 0, 'the middle itself');
  assert.equal(rotationAt(box, NaN, 50, false), 0);
  const handle = rotateHandlePoint(box, 90);
  near(handle.x, 120, 'handle x');
  near(handle.y, 50, 'handle y');
});

test('croppedBox and draggedCrop: what is still shown stays where it is, flipped pictures are cut on the side they show', () => {
  const box = { x: 100, y: 100, width: 200, height: 100 };
  // A quarter cut off the left: the box loses a quarter of its width on the left.
  assert.deepEqual(croppedBox(box, 0, undefined, { left: 0.25, top: 0, right: 0, bottom: 0 }), { x: 150, y: 100, width: 150, height: 100 });
  // Given back: the box as it was.
  const back = croppedBox({ x: 150, y: 100, width: 150, height: 100 }, 0, { left: 0.25, top: 0, right: 0, bottom: 0 }, undefined);
  assert.deepEqual(back, box);
  assert.deepEqual(uncroppedBox({ x: 150, y: 100, width: 150, height: 100 }, { left: 0.25, top: 0, right: 0, bottom: 0 }), box);
  // Flipped: the left of the picture is on the right of the box.
  assert.deepEqual(croppedBox(box, 0, undefined, { left: 0.25, top: 0, right: 0, bottom: 0 }, true), { x: 100, y: 100, width: 150, height: 100 });
  // The handle on the left dragged 50 pixels right cuts a quarter off; dragged beyond the picture it cuts nothing.
  assert.deepEqual(draggedCrop(box, 0, undefined, 'w', 50, 0, 4), { left: 0.25, top: 0, right: 0, bottom: 0 });
  assert.equal(draggedCrop(box, 0, undefined, 'w', -50, 0, 4), undefined);
  // So much that less than 4 pixels would be left: 4 are left.
  const most = draggedCrop(box, 0, undefined, 'se', -1000, -1000, 4)!;
  near(most.right, 1 - 4 / 200, 'right');
  near(most.bottom, 1 - 4 / 100, 'bottom');
  // Flipped, the handle on the left cuts the right of the picture.
  assert.deepEqual(draggedCrop(box, 0, undefined, 'w', 50, 0, 4, true), { left: 0, top: 0, right: 0.25, bottom: 0 });
  // Turned half way: the handle "n" is at the bottom on the screen; dragging it up (into the picture) cuts the top.
  assert.deepEqual(draggedCrop(box, 180, undefined, 'n', 0, -25, 4), { left: 0, top: 0.25, right: 0, bottom: 0 });
});

// ------------------------------------------------------------------ turn

test('pictureRotation: degrees as the file holds them, from 0 to under 360', () => {
  assert.equal(pictureRotation(90), 90);
  assert.equal(pictureRotation(-90), 270);
  assert.equal(pictureRotation(360), 0);
  assert.equal(pictureRotation(370.5), 10.5);
  assert.equal(pictureRotation(1e-9), 0);
  assert.equal(pictureRotation(359.9999999), 0);
  for (const bad of [NaN, Infinity, '90', undefined, 1e9]) assert.throws(() => pictureRotation(bad), RangeError);
});

test('an added picture is turned: its box stays, a quarter turn gives its anchor turned, and the file says so', async () => {
  const doc = await blank();
  apply(doc, addPicture(doc.wb, 0, png(4, 3), 'image/png', { kind: 'absolute', x: 200 * 9525, y: 100 * 9525, width: 200 * 9525, height: 100 * 9525 }));
  const anchor = doc.wb.sheets[0].drawings![0].anchor;
  apply(doc, rotatePicture(doc.wb, 0, 0, 30));
  assert.equal(pictureOf(doc.wb, 0).rotation, 30);
  assert.deepEqual(doc.wb.sheets[0].drawings![0].anchor, anchor, 'no quarter turn crossed: the anchor stays');
  assert.deepEqual(rotatePicture(doc.wb, 0, 0, 390), [], 'turned so already');
  apply(doc, rotatePicture(doc.wb, 0, 0, 90));
  assert.deepEqual(boxOf(doc.wb, 0), { x: 250, y: 50, width: 100, height: 200 });
  const saved = await save(doc);
  const xml = await partText(saved, 'xl/drawings/drawing1.xml');
  // Drawn in the box as it is not turned; the anchor is the box it takes up.
  assert.match(xml, /<a:xfrm rot="5400000"><a:off x="1905000" y="952500"\/><a:ext cx="1905000" cy="952500"\/><\/a:xfrm>/);
  assert.match(xml, /<xdr:absoluteAnchor><xdr:pos x="2381250" y="476250"\/><xdr:ext cx="952500" cy="1905000"\/>/);
  const again = await load(saved);
  assert.equal(pictureOf(again.wb, 0).rotation, 90);
  assert.deepEqual(boxOf(again.wb, 0), { x: 250, y: 50, width: 100, height: 200 });
  // Turned back: the box it was.
  apply(doc, rotatePicture(doc.wb, 0, 0, 0));
  assert.equal(pictureOf(doc.wb, 0).rotation, undefined);
  assert.deepEqual(boxOf(doc.wb, 0), { x: 200, y: 100, width: 200, height: 100 });
  if (python) assert.deepEqual(packageProblems({ 'turned.xlsx': saved }), []);
  if (python) {
    const outside = readOutside({ 'turned.xlsx': saved })['turned.xlsx'];
    assert.equal(outside.pictures[0].rot, '5400000');
    if (openpyxl) assert.deepEqual(outside.images, [1]);
  }
});

test('a picture of the file is turned: only its transform and its anchor change, everything else of the part keeps its bytes; undone, the file is as it was', async () => {
  const file = await pictureWorkbook();
  const doc = await load(file);
  apply(doc, rotatePicture(doc.wb, 0, 0, 45));
  const saved = await save(doc);
  assert.deepEqual(await changedParts(file, saved), ['xl/drawings/drawing1.xml']);
  const before = await partText(file, 'xl/drawings/drawing1.xml');
  const after = await partText(saved, 'xl/drawings/drawing1.xml');
  // The transform of the first picture has its turn; its anchor is the box turned (45 crosses the line); the rest is as it was.
  assert.equal((after.match(/rot="2700000"/g) || []).length, 1);
  assert.match(after, /<xdr:spPr><a:xfrm rot="2700000"><a:off x="0" y="0"\/><a:ext cx="0" cy="0"\/><\/a:xfrm>/);
  assert.match(after, /<a14:useLocalDpi xmlns:a14="http:\/\/schemas.microsoft.com\/office\/drawing\/2010\/main" val="0"\/>/);
  const tail = before.slice(before.indexOf('</xdr:twoCellAnchor>'));
  assert.ok(after.endsWith(tail), 'the other anchors keep their bytes');
  const again = await load(saved);
  assert.equal(pictureOf(again.wb, 0).rotation, 45);
  assert.deepEqual(sheetPictures(again.wb, 0)[0].anchor, sheetPictures(doc.wb, 0)[0].anchor);
  if (python) assert.deepEqual(packageProblems({ 'file-turned.xlsx': saved }), []);
  doc.history.undo();
  assert.ok(sameBytes(await save(doc), file), 'undone: byte for byte');
});

test('a picture of the file without a transform gets one of its turn alone; one turned in the file is turned back to none', async () => {
  const bare = pic('rIdI1', { name: 'Bare' }).replace(/<xdr:spPr>.*<\/xdr:spPr>/, '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>');
  const turned = pic('rIdI1', { name: 'Turned', xfrm: 'rot="1800000" flipH="1"' }).replace('cNvPr id="2"', 'cNvPr id="3"');
  const drawing =
    '<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="476250"/>' + bare + '<xdr:clientData/></xdr:absoluteAnchor>' +
    '<xdr:absoluteAnchor><xdr:pos x="1905000" y="0"/><xdr:ext cx="952500" cy="476250"/>' + turned + '<xdr:clientData/></xdr:absoluteAnchor>';
  const file = await pictureWorkbook({ drawing });
  const doc = await load(file);
  assert.equal(pictureOf(doc.wb, 1).rotation, 30);
  assert.equal(pictureOf(doc.wb, 1).flipH, true);
  apply(doc, rotatePicture(doc.wb, 0, 0, 10));
  apply(doc, rotatePicture(doc.wb, 0, 1, 0));
  const xml = await partText(await save(doc), 'xl/drawings/drawing1.xml');
  assert.match(xml, /<xdr:spPr><a:xfrm rot="600000"\/><a:prstGeom prst="rect">/);
  // The turn goes; the flip stays.
  assert.match(xml, /name="Turned".*<a:xfrm flipH="1"><a:off x="0" y="0"\/>/);
  const again = await load(await save(doc));
  assert.equal(pictureOf(again.wb, 0).rotation, 10);
  assert.equal(pictureOf(again.wb, 1).rotation, undefined);
  assert.equal(pictureOf(again.wb, 1).flipH, true);
  if (python) assert.deepEqual(packageProblems({ 'bare.xlsx': await save(doc) }), []);
});

// ------------------------------------------------------------------ crop

test('checkedCrop: shares from 0 to 1 in the steps of the file; nothing cut is no crop; one that leaves nothing is refused', () => {
  assert.equal(checkedCrop(undefined), undefined);
  assert.equal(checkedCrop({ left: 0, top: 0, right: 0, bottom: 0 }), undefined);
  assert.deepEqual(checkedCrop({ left: 0.123456789, top: 0, right: 0, bottom: 0.5 }), { left: 0.12346, top: 0, right: 0, bottom: 0.5 });
  assert.deepEqual(checkedCrop({ left: 0.2 } as never), { left: 0.2, top: 0, right: 0, bottom: 0 });
  for (const bad of [{ left: -0.1, top: 0, right: 0, bottom: 0 }, { left: 1, top: 0, right: 0, bottom: 0 }, { left: NaN, top: 0, right: 0, bottom: 0 }, { left: '0.1', top: 0, right: 0, bottom: 0 }, [0.1], 'crop']) {
    assert.throws(() => checkedCrop(bad), RangeError, JSON.stringify(bad));
  }
  assert.throws(() => checkedCrop({ left: 0.6, top: 0, right: 0.4, bottom: 0 }), /nothing/);
});

test('an added picture is cropped: the box loses what is cut, the file holds a:srcRect after the blip, and the crop is read back', async () => {
  const doc = await blank();
  apply(doc, addPicture(doc.wb, 0, png(4, 3), 'image/png', { kind: 'absolute', x: 0, y: 0, width: 200 * 9525, height: 100 * 9525 }));
  apply(doc, cropPicture(doc.wb, 0, 0, { left: 0.25, top: 0.1, right: 0, bottom: 0 }));
  const box = boxOf(doc.wb, 0);
  near(box.x, 50, 'x', 0.001);
  near(box.y, 10, 'y', 0.001);
  near(box.width, 150, 'width', 0.001);
  near(box.height, 90, 'height', 0.001);
  assert.deepEqual(cropPicture(doc.wb, 0, 0, { left: 0.25, top: 0.1, right: 0, bottom: 0 }), [], 'cropped so already');
  const saved = await save(doc);
  const xml = await partText(saved, 'xl/drawings/drawing1.xml');
  assert.match(xml, /r:embed="rId1"\/><a:srcRect l="25000" t="10000"\/><a:stretch>/);
  const again = await load(saved);
  assert.deepEqual(pictureOf(again.wb, 0).crop, { left: 0.25, top: 0.1, right: 0, bottom: 0 });
  // Given back: the whole picture in the box it had.
  apply(doc, cropPicture(doc.wb, 0, 0, undefined));
  const whole = boxOf(doc.wb, 0);
  near(whole.x, 0, 'x', 0.001);
  near(whole.width, 200, 'width', 0.001);
  if (python) {
    assert.deepEqual(packageProblems({ 'cropped.xlsx': saved }), []);
    const outside = readOutside({ 'cropped.xlsx': saved })['cropped.xlsx'];
    assert.deepEqual(outside.pictures[0].crop, { l: '25000', t: '10000' });
    if (openpyxl) assert.deepEqual(outside.images, [1]);
  }
});

test('a picture of the file is cropped and given back: a:srcRect is put after the blip and taken out again; the other bytes stay', async () => {
  const file = await pictureWorkbook();
  const doc = await load(file);
  apply(doc, cropPicture(doc.wb, 0, 0, { left: 0, top: 0, right: 0.5, bottom: 0 }));
  let xml = await partText(await save(doc), 'xl/drawings/drawing1.xml');
  // The blip of the first picture holds what the engine has no field for: the crop comes after it, untouched.
  assert.match(xml, /val="0"\/><\/a:ext><\/a:extLst><\/a:blip><a:srcRect r="50000"\/><a:stretch>/);
  assert.equal((xml.match(/srcRect/g) || []).length, 1);
  apply(doc, cropPicture(doc.wb, 0, 0, { left: 0.1, top: 0, right: 0.2, bottom: 0 }));
  xml = await partText(await save(doc), 'xl/drawings/drawing1.xml');
  assert.match(xml, /<a:srcRect l="10000" r="20000"\/>/);
  apply(doc, cropPicture(doc.wb, 0, 0, undefined));
  xml = await partText(await save(doc), 'xl/drawings/drawing1.xml');
  assert.ok(!/srcRect/.test(xml));
  if (python) assert.deepEqual(packageProblems({ 'file-cropped.xlsx': await save(doc) }), []);
  doc.history.undo();
  doc.history.undo();
  doc.history.undo();
  assert.ok(sameBytes(await save(doc), file), 'every crop undone: byte for byte');
});

test('a crop of the file is changed side by side: the sides it gives are set, the element stays where it is', async () => {
  const cropped = pic('rIdI1', { name: 'C', fill: '<a:srcRect l="10000" t="20000"/><a:stretch><a:fillRect/></a:stretch>' });
  const file = await pictureWorkbook({ drawing: '<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="476250"/>' + cropped + '<xdr:clientData/></xdr:absoluteAnchor>' });
  const doc = await load(file);
  assert.deepEqual(pictureOf(doc.wb, 0).crop, { left: 0.1, top: 0.2, right: 0, bottom: 0 });
  apply(doc, cropPicture(doc.wb, 0, 0, { left: 0.1, top: 0, right: 0, bottom: 0.3 }));
  const xml = await partText(await save(doc), 'xl/drawings/drawing1.xml');
  assert.match(xml, /<a:srcRect l="10000" b="30000"\/><a:stretch>/);
});

test('a turned picture is cropped in its own frame: its middle moves along its turn', async () => {
  const doc = await blank();
  apply(doc, addPicture(doc.wb, 0, png(4, 3), 'image/png', { kind: 'absolute', x: 200 * 9525, y: 200 * 9525, width: 200 * 9525, height: 100 * 9525 }));
  apply(doc, rotatePicture(doc.wb, 0, 0, 90));
  apply(doc, cropPicture(doc.wb, 0, 0, { left: 0.5, top: 0, right: 0, bottom: 0 }));
  // Turned a quarter to the right, the left of the picture is at the top: the box loses its upper half.
  const box = boxOf(doc.wb, 0);
  near(box.x, 250, 'x', 0.001);
  near(box.y, 250, 'y', 0.001);
  near(box.width, 100, 'width', 0.001);
  near(box.height, 100, 'height', 0.001);
});

// ------------------------------------------------------------------ alternative text

test('describePicture: the alternative text of a picture of the file changes in its descr alone, and goes', async () => {
  const file = await pictureWorkbook();
  const doc = await load(file);
  assert.equal(pictureOf(doc.wb, 0).description, 'Bir logo');
  assert.deepEqual(describePicture(doc.wb, 0, 0, 'Bir logo'), []);
  apply(doc, describePicture(doc.wb, 0, 0, 'Firma "logosu" & <imza>\nikinci satır'));
  let saved = await save(doc);
  assert.deepEqual(await changedParts(file, saved), ['xl/drawings/drawing1.xml']);
  let xml = await partText(saved, 'xl/drawings/drawing1.xml');
  assert.match(xml, /<xdr:cNvPr id="2" name="Logo" descr="Firma &quot;logosu&quot; &amp; &lt;imza>&#10;ikinci satır"\/>/);
  assert.equal(pictureOf((await load(saved)).wb, 0).description, 'Firma "logosu" & <imza>\nikinci satır');
  apply(doc, describePicture(doc.wb, 0, 1, 'İkinci görsel'));
  apply(doc, describePicture(doc.wb, 0, 0, ''));
  saved = await save(doc);
  xml = await partText(saved, 'xl/drawings/drawing1.xml');
  assert.match(xml, /<xdr:cNvPr id="2" name="Logo"\/>/);
  assert.match(xml, /name="İkinci" descr="İkinci görsel"/);
  if (python) {
    assert.deepEqual(packageProblems({ 'described.xlsx': saved }), []);
    assert.deepEqual(readOutside({ 'described.xlsx': saved })['described.xlsx'].pictures.map((item) => item.descr), [null, 'İkinci görsel', null]);
  }
  assert.throws(() => describePicture(doc.wb, 0, 0, 'x'.repeat(2001)), /2000/);
  assert.throws(() => describePicture(doc.wb, 0, 0, 'a\u0001b'), /character/);
  doc.history.undo();
  doc.history.undo();
  doc.history.undo();
  assert.ok(sameBytes(await save(doc), file));
});

// ------------------------------------------------------------------ copies

test('a picture of the file is copied with its bytes, and pasted beside it, then on another sheet and at a cell', async () => {
  const file = await pictureWorkbook();
  const doc = await load(file);
  apply(doc, rotatePicture(doc.wb, 0, 0, 30));
  apply(doc, cropPicture(doc.wb, 0, 0, { left: 0.1, top: 0, right: 0, bottom: 0 }));
  const copy = await copyPicture(doc.wb, 0, 0);
  assert.equal(copy.type, 'image/png');
  assert.ok(sameBytes(copy.bytes, png(4, 3)));
  assert.equal(copy.rotation, 30);
  assert.deepEqual(copy.crop, { left: 0.1, top: 0, right: 0, bottom: 0 });
  assert.equal(copy.description, 'Bir logo');
  const box = boxOf(doc.wb, 0);
  // Beside it: moved right and down, with its size, turn, crop, text and kind of anchor.
  const pasted = pastePicture(doc.wb, 0, copy);
  assert.equal(pasted.id, 3);
  apply(doc, pasted.ops);
  const placed = boxOf(doc.wb, 3);
  near(placed.x, box.x + PICTURE_COPY_OFFSET, 'x');
  near(placed.y, box.y + PICTURE_COPY_OFFSET, 'y');
  near(placed.width, box.width, 'width');
  const object = pictureOf(doc.wb, 3);
  assert.equal(object.name, 'Picture 4');
  assert.equal(object.rotation, 30);
  assert.equal(object.description, 'Bir logo');
  assert.equal(doc.wb.sheets[0].drawings![3].anchor.kind, 'twoCell');
  // A second paste beside the first copy.
  apply(doc, pastePicture(doc.wb, 0, copy).ops);
  near(boxOf(doc.wb, 4).x, box.x + 2 * PICTURE_COPY_OFFSET, 'x of the second');
  // On another sheet, at a cell.
  apply(doc, [{ type: 'insertSheet', sheet: 1, name: 'İkinci' }]);
  apply(doc, pastePicture(doc.wb, 1, copy, { row: 4, col: 2 }).ops);
  assert.deepEqual(doc.wb.sheets[1].drawings![0].anchor.kind === 'twoCell' && doc.wb.sheets[1].drawings![0].anchor.from, { col: 2, colOff: 0, row: 4, rowOff: 0 });
  const saved = await save(doc);
  const zip = ZipArchive.read(saved);
  // The bytes of the copies are those of the picture: one part for them on each sheet, the picture of the file stays in its own.
  const media = zip.names().filter((name) => name.startsWith('xl/media/')).sort();
  for (const name of media) assert.ok(sameBytes(await zip.read(name), png(4, 3)), name);
  assert.equal(media.length, 3);
  const again = await load(saved);
  assert.equal(sheetPictures(again.wb, 0).length, 5);
  assert.equal(sheetPictures(again.wb, 1).length, 1);
  assert.equal(pictureOf(again.wb, 3).rotation, 30);
  assert.deepEqual(pictureOf(again.wb, 0, 1).crop, { left: 0.1, top: 0, right: 0, bottom: 0 });
  if (python) {
    assert.deepEqual(packageProblems({ 'pasted.xlsx': saved }), []);
    const outside = readOutside({ 'pasted.xlsx': saved })['pasted.xlsx'];
    if (openpyxl) assert.deepEqual(outside.images, [5, 1]);
    assert.deepEqual(
      outside.pictures.map((item) => [item.name, item.rot]),
      [
        ['Logo', '1800000'],
        ['İkinci', null],
        ['Üçüncü', null],
        ['Picture 4', '1800000'],
        ['Picture 5', '1800000'],
        ['Picture 1', '1800000'],
      ],
    );
  }
});

test('a copy is pasted into another workbook: it names nothing of the one it came from', async () => {
  const from = await load(await pictureWorkbook());
  const copy = await copyPicture(from.wb, 0, 2);
  const to = await blank();
  apply(to, pastePicture(to.wb, 0, copy).ops);
  const [info] = sheetPictures(to.wb, 0);
  assert.equal(info.anchor.kind, 'absolute');
  assert.equal(info.name, 'Picture 1');
  const saved = await save(to);
  assert.ok(sameBytes(await ZipArchive.read(saved).read('xl/media/image1.png'), png(4, 3)));
  if (python) assert.deepEqual(packageProblems({ 'other.xlsx': saved }), []);
});

test('cutPicture: one step deletes it once the copy is made, and pasted is another step; both undone, the file is as it was', async () => {
  const file = await pictureWorkbook();
  const doc = await load(file);
  const cut = await cutPicture(doc.wb, 0, 1);
  apply(doc, cut.ops);
  assert.equal(sheetPictures(doc.wb, 0).length, 2);
  // Saved in between: the copy holds the bytes.
  await save(doc);
  const copy = cut.copy;
  assert.ok(sameBytes(copy.bytes, png(4, 3)));
  apply(doc, pastePicture(doc.wb, 0, copy, { row: 10, col: 1 }).ops);
  assert.equal(sheetPictures(doc.wb, 0).length, 3);
  doc.history.undo();
  doc.history.undo();
  assert.ok(sameBytes(await save(doc), file));
});

test('an added picture is copied from the model, and pasted with the same bytes: the writer writes them once', async () => {
  const doc = await blank();
  apply(doc, addPicture(doc.wb, 0, png(8, 8), 'image/png', { top: 0, left: 0, bottom: 3, right: 3 }, { description: 'Kaşe' }));
  const cut = await cutPicture(doc.wb, 0, 0);
  apply(doc, cut.ops);
  // The bytes left the model with the picture; the copy holds them.
  const copy = cut.copy;
  apply(doc, pastePicture(doc.wb, 0, copy).ops);
  apply(doc, pastePicture(doc.wb, 0, copy).ops);
  assert.equal(pictureOf(doc.wb, 0).data!.bytes, pictureOf(doc.wb, 1).data!.bytes);
  const saved = await save(doc);
  assert.deepEqual(ZipArchive.read(saved).names().filter((name) => name.startsWith('xl/media/')), ['xl/media/image1.png']);
  assert.equal(pictureOf((await load(saved)).wb, 0).description, 'Kaşe');
});

test('what cannot be copied is said at once; bytes that are no picture refuse the copy; a copy made elsewhere is checked as it is pasted', async () => {
  const linked = pic('rIdL', { name: 'L', link: true });
  const meta = pic('rIdE', { name: 'E' }).replace('cNvPr id="2"', 'cNvPr id="3"');
  const drawing =
    '<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="952500" cy="476250"/>' + linked + '<xdr:clientData/></xdr:absoluteAnchor>' +
    '<xdr:absoluteAnchor><xdr:pos x="0" y="952500"/><xdr:ext cx="952500" cy="476250"/>' + meta + '<xdr:clientData/></xdr:absoluteAnchor>';
  const bytes = await pictureWorkbook({ drawing });
  const zip = ZipArchive.read(bytes);
  const rels = new TextDecoder().decode(await zip.read('xl/drawings/_rels/drawing1.xml.rels'));
  zip.set(
    'xl/drawings/_rels/drawing1.xml.rels',
    new TextEncoder().encode(
      rels.replace(
        '</Relationships>',
        '<Relationship Id="rIdL" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.com/a.png" TargetMode="External"/><Relationship Id="rIdE" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.emf"/></Relationships>',
      ),
    ),
  );
  zip.set('xl/media/image2.emf', new Uint8Array([1, 0, 0, 0]));
  const types = new TextDecoder().decode(await zip.read('[Content_Types].xml'));
  zip.set('[Content_Types].xml', new TextEncoder().encode(types.replace('</Types>', '<Default Extension="emf" ContentType="image/x-emf"/></Types>')));
  const doc = await load(await zip.write());
  assert.match(pictureCopyProblem(doc.wb, 0, 0)!, /not in the file/);
  assert.match(pictureCopyProblem(doc.wb, 0, 1)!, /only PNG, JPEG, GIF and BMP/);
  assert.match(pictureCopyProblem(doc.wb, 0, 7)!, /No picture/);
  assert.throws(() => copyPicture(doc.wb, 0, 1), RangeError);
  // A part called .png that holds no picture: the copy is refused once it is read.
  const broken = ZipArchive.read(await pictureWorkbook());
  broken.set('xl/media/image1.png', new TextEncoder().encode('not a picture at all'));
  const bad = await load(await broken.write());
  await assert.rejects(copyPicture(bad.wb, 0, 0), /Cannot copy this picture: the bytes are not a PNG/);
  // A copy made by hand: its bytes are checked, and taken as a copy of their own.
  const target = await blank();
  const made = { bytes: png(2, 2), type: 'image/png', rect: { x: 0, y: 0, width: 40, height: 40 }, anchor: { kind: 'oneCell' as const } };
  const pasted = pastePicture(target.wb, 0, made as PictureCopy);
  made.bytes[30] ^= 0xff;
  apply(target, pasted.ops);
  assert.ok(sameBytes(pictureOf(target.wb, 0).data!.bytes, png(2, 2)));
  assert.equal(target.wb.sheets[0].drawings![0].anchor.kind, 'oneCell');
  assert.throws(() => pastePicture(target.wb, 0, { ...made, bytes: new TextEncoder().encode('no') } as PictureCopy), /not a PNG/);
  assert.throws(() => pastePicture(target.wb, 0, { ...made, crop: { left: 2, top: 0, right: 0, bottom: 0 } } as PictureCopy), RangeError);
  assert.throws(() => pastePicture(target.wb, 0, { ...made, rotation: 'x' } as unknown as PictureCopy), RangeError);
  assert.throws(() => pastePicture(target.wb, 0, 'picture' as unknown as PictureCopy), /not a picture that was copied/);
  assert.throws(() => pastePicture(target.wb, 0, made as PictureCopy, { row: -1, col: 0 }), /cell/);
});

// ------------------------------------------------------------------ refused

test('a protected sheet and a room refuse every new command, before anything is changed', async () => {
  const locked = await load(await pictureWorkbook({ protection: '<sheetProtection sheet="1" objects="1"/>' }));
  assert.throws(() => rotatePicture(locked.wb, 0, 0, 90), /protected/);
  assert.throws(() => cropPicture(locked.wb, 0, 0, { left: 0.1, top: 0, right: 0, bottom: 0 }), /protected/);
  assert.throws(() => describePicture(locked.wb, 0, 0, 'x'), /protected/);
  assert.throws(() => cutPicture(locked.wb, 0, 0), /protected/);
  const copy = await copyPicture(locked.wb, 0, 0);
  assert.throws(() => pastePicture(locked.wb, 0, copy), /protected/);
  const doc = await load(await pictureWorkbook());
  for (const ops of [rotatePicture(doc.wb, 0, 0, 90), cropPicture(doc.wb, 0, 0, { left: 0.1, top: 0, right: 0, bottom: 0 }), describePicture(doc.wb, 0, 0, 'x'), pastePicture(doc.wb, 0, copy).ops]) {
    assert.equal(ops.length, 1);
    assert.throws(
      () => encodeOp(doc.wb, ops[0]),
      (error: unknown) => error instanceof CollabUnsupportedError && error.what === 'picture',
    );
  }
});

test('each command is one step of the history, and undo gives back the picture it changed', async () => {
  const doc = await load(await pictureWorkbook());
  const before = JSON.stringify(doc.wb.sheets[0].drawings![0]);
  apply(doc, rotatePicture(doc.wb, 0, 0, 90));
  apply(doc, cropPicture(doc.wb, 0, 0, { left: 0.2, top: 0, right: 0, bottom: 0 }));
  apply(doc, describePicture(doc.wb, 0, 0, 'Yeni'));
  doc.history.undo();
  assert.equal(pictureOf(doc.wb, 0).description, 'Bir logo');
  doc.history.undo();
  assert.equal(pictureOf(doc.wb, 0).crop, undefined);
  doc.history.undo();
  assert.equal(JSON.stringify(doc.wb.sheets[0].drawings![0]), before);
});

test('rows inserted after a picture was turned move it from where it was turned; saved and read again, it is where the model says', async () => {
  const doc = await load(await pictureWorkbook());
  apply(doc, rotatePicture(doc.wb, 0, 1, 90));
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }]);
  const now = sheetPictures(doc.wb, 0)[1].anchor;
  const saved = await save(doc);
  assert.deepEqual(sheetPictures((await load(saved)).wb, 0)[1].anchor, now);
  if (python) assert.deepEqual(packageProblems({ 'rows.xlsx': saved }), []);
});

test('editDrawingPart: a patch in a part of other prefixes is written in its prefixes; a picture the anchor does not hold is an error', () => {
  const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const source = `<?xml version="1.0" encoding="UTF-8"?>\n<wsDr xmlns="${NS_XDR}" xmlns:d="${NS_A}"><absoluteAnchor><pos x="0" y="0"/><ext cx="9525" cy="9525"/><pic><nvPicPr><cNvPr id="7" name="P"/><cNvPicPr/></nvPicPr><blipFill><d:blip/><d:stretch/></blipFill><spPr/></pic><clientData/></absoluteAnchor><absoluteAnchor><pos x="0" y="0"/><ext cx="9525" cy="9525"/><sp><nvSpPr><cNvPr id="8" name="S"/><cNvSpPr/></nvSpPr><spPr/></sp><clientData/></absoluteAnchor></wsDr>`;
  const patched = new Map([[0, { rotation: 90, crop: { left: 0.5, top: 0, right: 0, bottom: 0 }, description: 'A & B' }]]);
  const out = new TextDecoder().decode(editDrawingPart(new TextEncoder().encode(source), { moved: new Map(), removed: new Set(), added: [], patched }).bytes);
  assert.match(out, /<cNvPr id="7" name="P" descr="A &amp; B"\/>/);
  assert.match(out, /<blipFill><d:blip\/><d:srcRect l="50000"\/><d:stretch\/><\/blipFill><spPr><d:xfrm rot="5400000"\/><\/spPr>/);
  // The anchor of the shape keeps its bytes.
  assert.ok(out.endsWith(source.slice(source.lastIndexOf('<absoluteAnchor>'))));
  assert.throws(() => editDrawingPart(new TextEncoder().encode(source), { moved: new Map(), removed: new Set(), added: [], patched: new Map([[1, { rotation: 10 }]]) }), /not a picture/);
});

test('the writer refuses a picture of the file that changed in what it does not write', async () => {
  const doc = await load(await pictureWorkbook());
  const sheet = doc.wb.sheets[0];
  const list = sheet.drawings!.slice();
  list[0] = { ...list[0], object: { ...(list[0].object as DrawingPicture), name: 'Another' }, from: list[0] };
  apply(doc, [{ type: 'setDrawings', sheet: 0, drawings: list }]);
  await assert.rejects(save(doc), /in what is not written/);
});
