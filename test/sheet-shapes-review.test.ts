// What the review of shapes found: text of a shape that the reader keeps only in part, a known
// shape whose text has more than one look, characters a file cannot hold dropped without a word,
// a shape of the file that lost its place in the file written anew, frames without a fill that take
// every click inside them, line breaks and paragraph marks a new text would lose, and connectors
// left bound to a shape that was deleted.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { createBlank, type DrawingShape } from '../src/index';
import { buildXlsx } from '../src/formats/xlsx/build';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import { addShape, removeShape, setShapeText, shapeEditProblem, sheetShapes } from '../src/ops/shapes';
import { SheetDrawings } from '../src/render/sheet-drawings';
import { ZipArchive } from '../src/zip';
import { apply, load, partText, save } from './edit-commands-fixtures';
import { pictureWorkbook } from './picture-fixtures';
import { outline, para, run, solid, sp, twoCell, txBody } from './sheet-drawing-fixtures';

const fileWith = (anchors: string) => pictureWorkbook({ drawing: anchors });
const shapeOf = (doc: { wb: { sheets: { drawings?: { object: unknown }[] }[] } }, id = 0) => doc.wb.sheets[0].drawings![id].object as DrawingShape;

test('B1: a body the reader keeps only in part (too many paragraphs, too much text) is not plain: its text is not replaced', async () => {
  let many = '';
  for (let i = 0; i < 2500; i++) many += para(run(`p${i}`));
  const long = para(run('a'.repeat(40000)));
  const file = await fileWith(twoCell([1, 1], [4, 6], sp({ body: txBody(many) })) + twoCell([6, 1], [8, 6], sp({ body: txBody(long) }).replace('cNvPr id="3"', 'cNvPr id="4"')));
  const doc = await load(file);
  for (const id of [0, 1]) {
    assert.equal(shapeOf(doc, id).plainText, undefined, `shape ${id}`);
    assert.match(shapeEditProblem(doc.wb, 0, id, 'text') || '', /Cannot change the text of this shape/);
    assert.throws(() => setShapeText(doc.wb, 0, id, 'x'), /Cannot change the text of this shape/);
  }
});

test('B1, review: text exactly as long as the reader reads, then one more paragraph, is not plain: that paragraph would be lost', async () => {
  const full = para(run('a'.repeat(32767))) + '<a:p/>';
  const doc = await load(await fileWith(twoCell([1, 1], [4, 6], sp({ body: txBody(full) }))));
  assert.equal(shapeOf(doc).plainText, undefined);
  assert.throws(() => setShapeText(doc.wb, 0, 0, 'x'), /Cannot change the text of this shape/);
});

test('B2: a known shape whose text has more than one look keeps its text: a new text would take the look of its first run only', async () => {
  const doc = await load(await createBlank('xlsx', { locale: 'tr' }));
  apply(doc, addShape(doc.wb, 0, 'rect', undefined, { text: 'Kutu' }));
  const shape = shapeOf(doc);
  const first = shape.text!.paragraphs[0].runs[0];
  shape.text!.paragraphs[0].runs = [{ ...first, text: 'Kalın ', bold: true }, { ...first, text: 'italik', italic: true, color: '#FF0000' }];
  const again = await load(await buildXlsx(doc.wb));
  assert.equal(shapeOf(again).known, true);
  assert.match(shapeEditProblem(again.wb, 0, 0, 'text') || '', /more than one look/);
  assert.throws(() => setShapeText(again.wb, 0, 0, 'Kalın italik!'), /more than one look/);
  assert.equal(sheetShapes(again.wb, 0)[0].textEditable, false);
});

test('B3: what a file cannot hold in the text, the name and the description of a shape is written as U+FFFD and counted; a name too long is cut and said', async () => {
  const doc = await load(await createBlank('xlsx', { locale: 'tr' }));
  apply(doc, addShape(doc.wb, 0, 'rect', undefined, { text: 'Kutu' }));
  const shape = shapeOf(doc);
  shape.text!.paragraphs[0].runs[0].text = 'A\u0001B';
  shape.name = 'Ad\u0002' + 'x'.repeat(300);
  shape.description = 'Aç\u0003';
  const report = { parts: [] as unknown[], warnings: [] as string[] };
  const built = await buildXlsx(doc.wb, { report } as never);
  const xml = new TextDecoder().decode(await ZipArchive.read(built).read('xl/drawings/drawing1.xml'));
  assert.match(xml, /<a:t>A�B<\/a:t>/);
  assert.match(xml, /descr="Aç�"/);
  assert.ok(report.warnings.some((warning) => /U\+FFFD \(3 times\)/.test(warning)), JSON.stringify(report.warnings));
  assert.ok(report.warnings.some((warning) => /name of a shape longer than 255 characters/.test(warning)), JSON.stringify(report.warnings));
});

test('B4: a shape of the file that is in no drawing the sheet was read with is not written anew from its model: the save is refused', async () => {
  const style = '<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>';
  const doc = await load(await fileWith(twoCell([1, 1], [4, 6], sp({ props: '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs></a:gsLst></a:gradFill>', style }))));
  const drawing = doc.wb.sheets[0].drawings![0];
  apply(doc, [{ type: 'setDrawings', sheet: 0, drawings: [{ ...drawing, object: { ...drawing.object } }] }]);
  await assert.rejects(save(doc), /neither in the file nor/);
});

test('B5: a shape without a fill takes a click near its outline, and inside its text; the cells inside an empty frame stay the cells', async () => {
  const frame = sp({ props: '<a:noFill/>' + outline('000000') });
  const labelled = sp({ props: '<a:noFill/>' + outline('000000'), body: txBody(para(run('Etiket')), '<a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/>') }).replace('cNvPr id="3"', 'cNvPr id="4"');
  const doc = await load(await fileWith(twoCell([0, 0], [8, 20], frame) + twoCell([10, 0], [14, 4], labelled)));
  const sheet = doc.wb.sheets[0];
  const layer = new SheetDrawings(sheet);
  const metrics = createSheetMetrics(sheet);
  const shapes = (object: { type: string }) => object.type === 'shape';
  assert.equal(layer.hit(metrics, 4 * 64 + 10, 10 * 20 + 5, shapes), -1, 'the middle of an empty frame');
  assert.equal(layer.hit(metrics, 1, 10 * 20 + 5, shapes), 0, 'on its left side');
  assert.equal(layer.hit(metrics, 8 * 64 - 2, 10 * 20 + 5, shapes), 0, 'on its right side');
  assert.equal(layer.hit(metrics, 12 * 64, 2 * 20, shapes), 1, 'inside the text of a frame with text');
  // A filled shape takes a click anywhere inside, as before.
  const filled = await load(await fileWith(twoCell([0, 0], [8, 20], sp({ props: solid('FF0000') }))));
  const filledLayer = new SheetDrawings(filled.wb.sheets[0]);
  assert.equal(filledLayer.hit(createSheetMetrics(filled.wb.sheets[0]), 4 * 64 + 10, 10 * 20 + 5, shapes), 0);
});

test('B6: a body with line breaks, or with paragraphs that end in other marks, is not plain: a new text would lose them', async () => {
  const breaks = txBody(para(run('A') + '<a:br><a:rPr lang="en-US"/></a:br>' + run('B')));
  const marks = txBody(para(run('A') + '<a:endParaRPr lang="en-US" sz="1100"/>') + para(run('C') + '<a:endParaRPr lang="en-US" sz="2400"/>'));
  const doc = await load(await fileWith(twoCell([1, 1], [4, 6], sp({ body: breaks })) + twoCell([6, 1], [8, 6], sp({ body: marks }).replace('cNvPr id="3"', 'cNvPr id="4"'))));
  for (const id of [0, 1]) {
    assert.equal(shapeOf(doc, id).plainText, undefined, `shape ${id}`);
    assert.throws(() => setShapeText(doc.wb, 0, id, 'x'), /Cannot change the text of this shape/);
  }
});

test('B7: a connector bound to a shape that is deleted is bound to it no longer; one bound to a shape that stays keeps its binding', async () => {
  const box = (id: number, name: string) => sp({ name, props: solid('FF0000') }).replace('cNvPr id="3"', `cNvPr id="${id}"`);
  const connector =
    '<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="9" name="Bağlayıcı"/><xdr:cNvCxnSpPr><a:stCxn id="3" idx="3"/><a:endCxn id="4" idx="1"/></xdr:cNvCxnSpPr></xdr:nvCxnSpPr>' +
    `<xdr:spPr><a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>${outline('000000')}</xdr:spPr></xdr:cxnSp>`;
  const file = await fileWith(twoCell([1, 1], [2, 2], box(3, 'Bir')) + twoCell([5, 1], [6, 2], box(4, 'İki')) + twoCell([2, 1], [5, 2], connector));
  const doc = await load(file);
  apply(doc, removeShape(doc.wb, 0, 0));
  const xml = await partText(await save(doc), 'xl/drawings/drawing1.xml');
  assert.equal(xml.includes('<a:stCxn'), false, 'the end bound to the deleted shape');
  assert.ok(xml.includes('<a:endCxn id="4" idx="1"/>'), 'the end bound to the shape that stays');
  assert.ok(xml.includes('<xdr:cNvCxnSpPr><a:endCxn'), xml);
});
