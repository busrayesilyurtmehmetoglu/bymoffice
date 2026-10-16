// What the review of turning, cropping, describing and cutting pictures found: a picture the
// file holds in a wrapper of alternatives, a cut whose copy cannot be made, a turn and a turn back
// at the edge of the sheet, alternative texts as long as the reader takes, crops that would leave
// less than the least size, and the handle that turns a picture on a touch screen. The commands
// are reached through the module, so that each test stands on its own.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { createBlank, ZipArchive, type DrawingPicture, type Workbook } from '../src/index';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import * as ops from '../src/ops/pictures';
import * as frame from '../src/render/picture-frame';
import { anchorRect, drawingAnchorNow } from '../src/render/sheet-drawings';
import { apply, changedParts, hasPython, load, packageProblems, save, sameBytes } from './edit-commands-fixtures';
import { png, pngHeader, pictureWorkbook } from './picture-fixtures';
import { pic, twoCell } from './sheet-drawing-fixtures';

const python = hasPython(['xml.etree.ElementTree', 'zipfile']);

async function blank() {
  return load(await createBlank('xlsx', { locale: 'tr' }));
}

const pictureOf = (wb: Workbook, id: number, sheet = 0) => wb.sheets[sheet].drawings![id].object as DrawingPicture;
const boxOf = (wb: Workbook, id: number, sheet = 0) => {
  const s = wb.sheets[sheet];
  return anchorRect(drawingAnchorNow(s.drawings![id], s.structuralEdits), createSheetMetrics(s));
};
const twoCellOf = (inner: string) => twoCell([1, 1], [4, 6], inner);
const inAlternatives = (inner: string) => '<mc:AlternateContent><mc:Choice Requires="a14">' + inner + '</mc:Choice><mc:Fallback>' + inner.replace('cNvPr id="2"', 'cNvPr id="3"') + '</mc:Fallback></mc:AlternateContent>';

test('a picture the file holds in a wrapper of alternatives is not turned, cropped or described: refused before anything is changed, and the workbook is still saved', async () => {
  // The picture inside its anchor in two forms, and the anchor itself in two forms.
  for (const drawing of [twoCellOf(inAlternatives(pic('rIdI1', { name: 'AC' }))), inAlternatives(twoCellOf(pic('rIdI1', { name: 'AC' })))]) {
    const file = await pictureWorkbook({ drawing });
    const doc = await load(file);
    assert.equal(typeof ops.pictureEditProblem, 'function');
    assert.match(ops.pictureEditProblem(doc.wb, 0, 0)!, /another form/);
    assert.throws(() => ops.rotatePicture(doc.wb, 0, 0, 30), /another form/);
    assert.throws(() => ops.cropPicture(doc.wb, 0, 0, { left: 0.1, top: 0, right: 0, bottom: 0 }), /another form/);
    assert.throws(() => ops.describePicture(doc.wb, 0, 0, 'x'), /another form/);
    // It is still moved (its anchor alone changes), and the workbook is saved.
    apply(doc, ops.movePicture(doc.wb, 0, 0, { top: 3, left: 3, bottom: 6, right: 6 }));
    const saved = await save(doc);
    assert.deepEqual(await changedParts(file, saved), ['xl/drawings/drawing1.xml']);
    if (python) assert.deepEqual(packageProblems({ 'wrapped.xlsx': saved }), []);
  }
  const plain = await load(await pictureWorkbook());
  assert.equal(ops.pictureEditProblem(plain.wb, 0, 0), undefined);
});

test('cutPicture: a picture whose copy cannot be made is not deleted; the deletion comes with the copy', async () => {
  const zip = ZipArchive.read(await pictureWorkbook());
  zip.set('xl/media/image1.png', pngHeader(20000, 20000));
  const doc = await load(await zip.write());
  const cut = ops.cutPicture(doc.wb, 0, 0) as unknown;
  assert.ok(cut instanceof Promise, 'the cut waits for its copy');
  await assert.rejects(cut as Promise<unknown>, /Cannot copy this picture/);
  assert.equal(doc.wb.sheets[0].drawings!.length, 3);
  // A picture that can be copied: the operations come with the copy, and delete the picture.
  const good = await load(await pictureWorkbook());
  const made = await ops.cutPicture(good.wb, 0, 1);
  assert.ok(sameBytes(made.copy.bytes, png(4, 3)));
  apply(good, made.ops);
  assert.equal(good.wb.sheets[0].drawings!.length, 2);
  // Deleted meanwhile: nothing is cut.
  const gone = await load(await pictureWorkbook());
  const late = ops.cutPicture(gone.wb, 0, 2);
  apply(gone, ops.removePicture(gone.wb, 0, 2));
  await assert.rejects(late, /changed or deleted/);
});

test('turned and turned back at the edge of the sheet, a picture stands where it stood', async () => {
  const doc = await blank();
  const start = { kind: 'absolute' as const, x: 0, y: 0, width: 400 * 9525, height: 50 * 9525 };
  apply(doc, ops.addPicture(doc.wb, 0, png(4, 3), 'image/png', start));
  apply(doc, ops.rotatePicture(doc.wb, 0, 0, 90));
  // Kept on the sheet while it is turned.
  assert.ok(boxOf(doc.wb, 0).y >= 0);
  apply(doc, ops.rotatePicture(doc.wb, 0, 0, 0));
  assert.deepEqual(doc.wb.sheets[0].drawings![0].anchor, start);
  // A cell anchor too, by way of another turn across the line.
  const cells = await blank();
  apply(cells, ops.addPicture(cells.wb, 0, png(4, 3), 'image/png', { top: 0, left: 0, bottom: 2, right: 6 }));
  const anchor = cells.wb.sheets[0].drawings![0].anchor;
  apply(cells, ops.rotatePicture(cells.wb, 0, 0, 270));
  apply(cells, ops.rotatePicture(cells.wb, 0, 0, 10));
  assert.deepEqual(cells.wb.sheets[0].drawings![0].anchor, anchor);
  // Moved in between: turned back from where it is.
  apply(doc, ops.rotatePicture(doc.wb, 0, 0, 90));
  apply(doc, ops.movePicture(doc.wb, 0, 0, { kind: 'absolute', x: 500 * 9525, y: 500 * 9525, width: 50 * 9525, height: 400 * 9525 }));
  apply(doc, ops.rotatePicture(doc.wb, 0, 0, 0));
  assert.deepEqual(boxOf(doc.wb, 0), { x: 325, y: 675, width: 400, height: 50 });
  // Rows inserted in between: turned back from where it is too.
  apply(cells, ops.rotatePicture(cells.wb, 0, 0, 90));
  apply(cells, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 3 }]);
  const before = boxOf(cells.wb, 0);
  apply(cells, ops.rotatePicture(cells.wb, 0, 0, 0));
  const after = boxOf(cells.wb, 0);
  assert.ok(Math.abs(after.x + after.width / 2 - (before.x + before.width / 2)) < 1 && Math.abs(after.y + after.height / 2 - (before.y + before.height / 2)) < 1);
});

test('the alternative text of the file is kept whole up to the limit the reader takes, and a copy takes it as it is', async () => {
  const long = 'x'.repeat(1500) + '😀';
  const doc = await load(await pictureWorkbook({ drawing: twoCellOf(pic('rIdI1', { name: 'L', descr: long })) }));
  assert.equal(ops.MAX_PICTURE_TEXT, 2000);
  assert.equal(pictureOf(doc.wb, 0).description, long);
  assert.equal((await ops.copyPicture(doc.wb, 0, 0)).description, long);
  // A small change of a long text is no text too long.
  apply(doc, ops.describePicture(doc.wb, 0, 0, long + '!'));
  assert.throws(() => ops.describePicture(doc.wb, 0, 0, 'y'.repeat(2001)), /2000/);
  // Longer than the limit in the file: cut without splitting a character in two.
  const longer = await load(await pictureWorkbook({ drawing: twoCellOf(pic('rIdI1', { name: 'L', descr: 'x'.repeat(1999) + '😀' })) }));
  assert.equal(pictureOf(longer.wb, 0).description, 'x'.repeat(1999));
});

test('a crop that would leave less than the least size of a picture is refused, as the handles refuse it', async () => {
  const doc = await blank();
  apply(doc, ops.addPicture(doc.wb, 0, png(4, 3), 'image/png', { kind: 'absolute', x: 0, y: 0, width: 100 * 9525, height: 100 * 9525 }));
  assert.throws(() => ops.cropPicture(doc.wb, 0, 0, { left: 0.5, top: 0, right: 0.4989, bottom: 0 }), /4 pixels/);
  apply(doc, ops.cropPicture(doc.wb, 0, 0, { left: 0.5, top: 0, right: 0.46, bottom: 0 }));
  assert.ok(Math.abs(boxOf(doc.wb, 0).width - 4) < 0.01);
});

test('rotateHandlePoint: on a touch screen the handle that turns stands further up, clear of the handle of the top side', () => {
  const box = { x: 0, y: 100, width: 100, height: 50 };
  const reach = 16;
  const coarse = (frame as unknown as { ROTATE_HANDLE_GAP_COARSE?: number }).ROTATE_HANDLE_GAP_COARSE;
  assert.equal(typeof coarse, 'number');
  const knob = frame.rotateHandlePoint(box, 0, coarse as number);
  assert.ok(box.y - knob.y - reach > reach, 'the bands of the two handles do not meet');
  assert.equal(frame.rotateHandlePoint(box, 0).y, box.y - frame.ROTATE_HANDLE_GAP);
});
