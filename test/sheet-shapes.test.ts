// Shapes of a sheet added, moved, sized, turned, formatted, given text, copied and deleted
// (ops/shapes.ts), and what the writer makes of them: a shape that is added is written from its
// model (shape-write.ts) and read back as "known"; of a shape of the file only its anchor, its turn
// (`a:xfrm rot`) and the text of a plain body change, everything else of the part keeps its bytes;
// what is undone leaves the file byte for byte as it was. Saved packages are read back by the
// engine, by test/tools/package_check.py, and by ElementTree and openpyxl.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { encodeOp } from '../src/collab/wire';
import { CollabUnsupportedError } from '../src/collab/types';
import { createBlank, ZipArchive, type DrawingShape, type Workbook } from '../src/index';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import { buildXlsx } from '../src/formats/xlsx/build';
import { sameShapeElement, shapeElementXml } from '../src/formats/xlsx/shape-write';
import {
  addShape,
  checkedShapeText,
  copyShape,
  cutShape,
  formatShape,
  moveShape,
  moveShapeTo,
  pasteShape,
  removeShape,
  rotateShape,
  setShapeText,
  shapeCopyProblem,
  shapeEditProblem,
  shapeFormatProblem,
  shapesProblem,
  sheetShapes,
  SHAPE_COPY_OFFSET,
  SHAPE_KINDS,
  type ShapeCopy,
} from '../src/ops/shapes';
import { anchorRect, drawingAnchorNow, SheetDrawings } from '../src/render/sheet-drawings';
import { parseXml } from '../src/xml';
import { apply, changedParts, hasPython, load, packageProblems, partText, runPython, save, sameBytes } from './edit-commands-fixtures';
import { pictureWorkbook } from './picture-fixtures';
import { para, run, solid, outline, sp, txBody, twoCell, xfrm } from './sheet-drawing-fixtures';

const python = hasPython(['xml.etree.ElementTree', 'zipfile']);
const openpyxl = hasPython(['openpyxl']);

async function blank() {
  return load(await createBlank('xlsx', { locale: 'tr' }));
}

const shapeOf = (wb: Workbook, id: number, sheet = 0) => wb.sheets[sheet].drawings![id].object as DrawingShape;
const boxOf = (wb: Workbook, id: number, sheet = 0) => {
  const s = wb.sheets[sheet];
  return anchorRect(drawingAnchorNow(s.drawings![id], s.structuralEdits), createSheetMetrics(s));
};
const at = (x: number, y: number, width: number, height: number) => ({ kind: 'absolute' as const, x: x * 9525, y: y * 9525, width: width * 9525, height: height * 9525 });

/** A workbook whose drawing holds the shapes given (inner XML of the anchors), with the picture fixture's parts around it. */
const fileWith = (anchors: string) => pictureWorkbook({ drawing: anchors });

/** A shape as Excel writes one: a style of the theme beside its own properties, and a body in one look. */
const excelShape = (options: { name?: string; id?: number; body?: string; extra?: string } = {}) =>
  sp({
    name: options.name ?? 'Rectangle 1',
    xfrm: xfrm(0, 0, 0, 0),
    props: solid('4472C4') + outline('2F528F'),
    style: '<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>',
    body:
      options.body ??
      txBody(para(run('Merhaba', 'sz="1400"', solid('FF0000')) + '<a:endParaRPr lang="tr-TR" sz="1400"/>', '<a:pPr algn="ctr"/>'), '<a:bodyPr vertOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/>'),
  }).replace('cNvPr id="3"', `cNvPr id="${options.id ?? 3}"`) + (options.extra ?? '');

/** What ElementTree finds of the shapes of every drawing part of each file: element, id, preset, fill, text, line ends; and whether openpyxl opens the file. */
function readOutside(files: { [name: string]: Uint8Array }): { [name: string]: { opens: boolean | null; shapes: { kind: string; id: string; name: string; prst: string | null; fill: string | null; text: string; head: string | null; tail: string | null; rot: string | null; txBox: string | null }[]; ids: string[] } } {
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
    shapes = []
    ids = []
    for part in sorted(n for n in z.namelist() if n.startswith('xl/drawings/') and n.endswith('.xml')):
        root = ET.fromstring(z.read(part))
        for el in root.iter():
            if el.tag == XDR + 'cNvPr':
                ids.append(el.get('id'))
        for el in root.iter():
            if el.tag not in (XDR + 'sp', XDR + 'cxnSp'):
                continue
            kind = 'sp' if el.tag == XDR + 'sp' else 'cxnSp'
            nv = el.find(XDR + ('nvSpPr' if kind == 'sp' else 'nvCxnSpPr'))
            props = nv.find(XDR + 'cNvPr')
            flags = nv.find(XDR + 'cNvSpPr')
            geom = el.find(XDR + 'spPr/' + A + 'prstGeom')
            fill = el.find(XDR + 'spPr/' + A + 'solidFill/' + A + 'srgbClr')
            head = el.find(XDR + 'spPr/' + A + 'ln/' + A + 'headEnd')
            tail = el.find(XDR + 'spPr/' + A + 'ln/' + A + 'tailEnd')
            transform = el.find(XDR + 'spPr/' + A + 'xfrm')
            lines = []
            body = el.find(XDR + 'txBody')
            if body is not None:
                for p in body.findall(A + 'p'):
                    lines.append(''.join(t.text or '' for t in p.iter(A + 't')))
            shapes.append({
                'kind': kind, 'id': props.get('id'), 'name': props.get('name'),
                'prst': geom.get('prst') if geom is not None else None,
                'fill': fill.get('val') if fill is not None else None,
                'text': '\n'.join(lines),
                'head': head.get('type') if head is not None else None,
                'tail': tail.get('type') if tail is not None else None,
                'rot': transform.get('rot') if transform is not None else None,
                'txBox': flags.get('txBox') if flags is not None else None,
            })
    opens = None
    if openpyxl is not None:
        try:
            openpyxl.load_workbook(path)
            opens = True
        except Exception:
            opens = False
    out[name] = {'opens': opens, 'shapes': shapes, 'ids': ids}
print(json.dumps(out))
`,
  );
}

// ------------------------------------------------------------------ adding

test('every kind of shape is added with the model its kind starts with: a line flat, an arrow with its ends, a text box marked', async () => {
  const doc = await blank();
  for (const kind of SHAPE_KINDS) apply(doc, addShape(doc.wb, 0, kind));
  const shapes = doc.wb.sheets[0].drawings!.map((drawing) => drawing.object as DrawingShape);
  assert.deepEqual(
    shapes.map((shape) => [shape.geometry, shape.connector === true]),
    [['rect', false], ['roundRect', false], ['ellipse', false], ['line', true], ['straightConnector1', true], ['straightConnector1', true], ['rect', false]],
  );
  assert.equal(shapes[4].line!.tail, 'triangle');
  assert.equal(shapes[4].line!.head, undefined);
  assert.equal(shapes[5].line!.head, 'triangle');
  assert.equal(shapes[6].textBox, true);
  assert.equal(shapes[3].fill, undefined, 'a line has no fill');
  assert.deepEqual(shapes.map((shape) => shape.name), ['Rectangle 1', 'Rectangle: Rounded Corners 2', 'Oval 3', 'Straight Connector 4', 'Straight Arrow Connector 5', 'Straight Arrow Connector 6', 'TextBox 7']);
  // At A1 in the default size; a line as wide and flat.
  assert.deepEqual(boxOf(doc.wb, 0), { x: 0, y: 0, width: 160, height: 80 });
  assert.deepEqual(boxOf(doc.wb, 3), { x: 0, y: 0, width: 160, height: 0 });
  assert.equal(doc.wb.sheets[0].drawings![0].anchor.kind, 'twoCell');
  assert.equal(doc.wb.sheets[0].drawings![0].anchor.editAs, undefined, 'it moves and sizes with its cells');
});

test('added shapes are written from their models, read back as the same models and as known; every reader opens the file', async () => {
  const doc = await blank();
  for (const kind of SHAPE_KINDS) apply(doc, addShape(doc.wb, 0, kind, undefined, kind === 'line' || kind === 'arrow' || kind === 'doubleArrow' ? {} : { text: kind === 'textBox' ? 'Bir\nİki & <üç>' : 'Kutu' }));
  apply(doc, moveShapeTo(doc.wb, 0, 3, { x: 10, y: 200, width: 300, height: 0 }));
  const saved = await save(doc);
  const again = await load(saved);
  const read = again.wb.sheets[0].drawings!;
  assert.equal(read.length, SHAPE_KINDS.length);
  for (let i = 0; i < read.length; i++) {
    const shape = read[i].object as DrawingShape;
    assert.equal(shape.known, true, `shape ${i} is known`);
    const made = shapeOf(doc.wb, i);
    assert.equal(shape.geometry, made.geometry);
    assert.equal(shape.connector === true, made.connector === true);
    assert.deepEqual(shape.line, made.line);
    assert.equal(shape.fill, made.fill);
    assert.deepEqual(shape.text, made.text);
    assert.equal(shape.name, made.name);
  }
  assert.deepEqual(boxOf(again.wb, 3), { x: 10, y: 200, width: 300, height: 0 }, 'a flat line stays flat');
  const xml = await partText(saved, 'xl/drawings/drawing1.xml');
  // Excel's elements: a line is a connector, the text box says so.
  assert.equal((xml.match(/<xdr:cxnSp /g) || []).length, 3);
  assert.equal((xml.match(/<xdr:sp /g) || []).length, 4);
  assert.match(xml, /<a:t>İki &amp; &lt;üç&gt;<\/a:t>/);
  if (python) {
    assert.deepEqual(packageProblems({ 'shapes.xlsx': saved }), []);
    const outside = readOutside({ 'shapes.xlsx': saved })['shapes.xlsx'];
    assert.deepEqual(outside.shapes.map((shape) => shape.prst), ['rect', 'roundRect', 'ellipse', 'line', 'straightConnector1', 'straightConnector1', 'rect']);
    assert.deepEqual(outside.shapes.map((shape) => shape.kind), ['sp', 'sp', 'sp', 'cxnSp', 'cxnSp', 'cxnSp', 'sp']);
    assert.equal(outside.shapes[6].txBox, '1');
    assert.equal(outside.shapes[6].text, 'Bir\nİki & <üç>');
    assert.equal(outside.shapes[4].tail, 'triangle');
    assert.equal(outside.shapes[5].head, 'triangle');
    assert.equal(new Set(outside.ids).size, outside.ids.length, 'the ids of the drawing differ');
    if (openpyxl) assert.equal(outside.opens, true);
  }
});

test('a shape added to a drawing of the file goes into it: its id is one no other object has, the other anchors keep their bytes; undone, the file is as it was', async () => {
  const file = await pictureWorkbook();
  const doc = await load(file);
  apply(doc, addShape(doc.wb, 0, 'ellipse', { top: 8, left: 1, bottom: 12, right: 3 }));
  const saved = await save(doc);
  assert.deepEqual(await changedParts(file, saved), ['xl/drawings/drawing1.xml']);
  const before = await partText(file, 'xl/drawings/drawing1.xml');
  const after = await partText(saved, 'xl/drawings/drawing1.xml');
  assert.ok(after.startsWith(before.slice(0, before.lastIndexOf('</xdr:wsDr>'))), 'what was there keeps its bytes');
  assert.match(after, /<xdr:cNvPr id="5" name="Oval 1"\/>/);
  if (python) {
    assert.deepEqual(packageProblems({ 'added.xlsx': saved }), []);
    const outside = readOutside({ 'added.xlsx': saved })['added.xlsx'];
    assert.equal(new Set(outside.ids).size, outside.ids.length);
  }
  doc.history.undo();
  assert.ok(sameBytes(await save(doc), file), 'undone: byte for byte');
});

test('a sheet without a drawing gets one for its first shape; undone, the file is as it was', async () => {
  const doc = await blank();
  const file = await save(doc);
  const again = await load(file);
  apply(again, addShape(again.wb, 0, 'rect'));
  const saved = await save(again);
  const names = ZipArchive.read(saved).names();
  assert.ok(names.includes('xl/drawings/drawing1.xml'));
  assert.ok(names.includes('xl/drawings/_rels/drawing1.xml.rels'));
  assert.match(await partText(saved, '[Content_Types].xml'), /drawing\+xml/);
  assert.match(await partText(saved, 'xl/worksheets/sheet1.xml'), /<drawing r:id="/);
  if (python) assert.deepEqual(packageProblems({ 'first.xlsx': saved }), []);
  again.history.undo();
  assert.ok(sameBytes(await save(again), file), 'undone: byte for byte');
});

// ------------------------------------------------------------------ changing added shapes

test('an added shape is moved, sized, turned (a quarter turn gives its anchor turned) and deleted, each one undo step', async () => {
  const doc = await blank();
  apply(doc, addShape(doc.wb, 0, 'rect', at(100, 100, 200, 100)));
  apply(doc, moveShapeTo(doc.wb, 0, 0, { x: 120, y: 140, width: 220, height: 60 }));
  assert.deepEqual(boxOf(doc.wb, 0), { x: 120, y: 140, width: 220, height: 60 });
  apply(doc, rotateShape(doc.wb, 0, 0, 30));
  assert.equal(shapeOf(doc.wb, 0).rotation, 30);
  assert.deepEqual(boxOf(doc.wb, 0), { x: 120, y: 140, width: 220, height: 60 }, 'no quarter crossed: the anchor stays');
  apply(doc, rotateShape(doc.wb, 0, 0, 90));
  assert.deepEqual(boxOf(doc.wb, 0), { x: 200, y: 60, width: 60, height: 220 });
  const saved = await save(doc);
  const xml = await partText(saved, 'xl/drawings/drawing1.xml');
  // Drawn in the box as it is not turned; the anchor is the box it takes up.
  assert.match(xml, /<a:xfrm rot="5400000"><a:off x="1143000" y="1333500"\/><a:ext cx="2095500" cy="571500"\/><\/a:xfrm>/);
  const again = await load(saved);
  assert.equal(shapeOf(again.wb, 0).rotation, 90);
  assert.equal(shapeOf(again.wb, 0).known, true);
  doc.history.undo();
  assert.deepEqual(boxOf(doc.wb, 0), { x: 120, y: 140, width: 220, height: 60 });
  apply(doc, removeShape(doc.wb, 0, 0));
  assert.equal(doc.wb.sheets[0].drawings, undefined);
  doc.history.undo();
  assert.equal(doc.wb.sheets[0].drawings!.length, 1);
  assert.deepEqual(rotateShape(doc.wb, 0, 0, 30), [], 'turned so already');
  assert.throws(() => moveShape(doc.wb, 0, 0, { kind: 'oneCell', from: { col: 0, colOff: 0, row: 0, rowOff: 0 }, width: 1, height: 1 }), /keeps the kind of its anchor/);
  assert.throws(() => rotateShape(doc.wb, 0, 0, NaN), RangeError);
  assert.throws(() => removeShape(doc.wb, 0, 7), /No shape at index 7/);
});

test('the fill, the outline and the look of the text of an added shape are changed; the file has them and they are read back', async () => {
  const doc = await blank();
  apply(doc, addShape(doc.wb, 0, 'rect', undefined, { text: 'A\nB' }));
  apply(doc, formatShape(doc.wb, 0, 0, { fill: '#ff0000', lineColor: '#00FF00', lineWidth: 3, textColor: '#0000FF', textSize: 18, bold: true, align: 'right' }));
  const shape = shapeOf(doc.wb, 0);
  assert.equal(shape.fill, '#FF0000');
  assert.deepEqual(shape.line, { color: '#00FF00', width: 4 });
  assert.deepEqual(shape.text!.paragraphs.map((paragraph) => [paragraph.align, paragraph.runs.map((r) => [r.text, r.color, r.size, r.bold])]), [
    ['right', [['A', '#0000FF', 18, true]]],
    ['right', [['B', '#0000FF', 18, true]]],
  ]);
  assert.deepEqual(formatShape(doc.wb, 0, 0, { fill: '#FF0000' }), [], 'so already');
  apply(doc, formatShape(doc.wb, 0, 0, { lineColor: null }));
  assert.equal(shapeOf(doc.wb, 0).line, undefined);
  const saved = await save(doc);
  const again = await load(saved);
  const read = shapeOf(again.wb, 0);
  assert.equal(read.known, true);
  assert.equal(read.fill, '#FF0000');
  assert.equal(read.line, undefined);
  assert.deepEqual(read.text, shapeOf(doc.wb, 0).text);
  // Typed into it later, text gets the look it was given.
  apply(again, setShapeText(again.wb, 0, 0, 'Yeni'));
  assert.deepEqual(shapeOf(again.wb, 0).text!.paragraphs[0].runs.map((r) => [r.text, r.color, r.size, r.bold]), [['Yeni', '#0000FF', 18, true]]);
  if (python) assert.deepEqual(packageProblems({ 'formatted.xlsx': saved }), []);
  // What a format cannot be.
  assert.throws(() => formatShape(doc.wb, 0, 0, { fill: 'red' }), /colour such as/);
  assert.throws(() => formatShape(doc.wb, 0, 0, { lineWidth: 100 }), /points from/);
  assert.throws(() => formatShape(doc.wb, 0, 0, { textSize: 0 }), /points from/);
  apply(doc, addShape(doc.wb, 0, 'line'));
  assert.throws(() => formatShape(doc.wb, 0, 1, { fill: '#FFFFFF' }), /A line has no fill/);
  assert.throws(() => formatShape(doc.wb, 0, 1, { lineColor: null }), /A line has an outline/);
  assert.throws(() => setShapeText(doc.wb, 0, 1, 'x'), /A line has no text/);
});

test('the text of a shape: lines are paragraphs, empty takes it away; what a file cannot hold or is too long is refused', async () => {
  const doc = await blank();
  apply(doc, addShape(doc.wb, 0, 'textBox'));
  apply(doc, setShapeText(doc.wb, 0, 0, 'bir\r\niki\n\ndört'));
  assert.deepEqual(sheetShapes(doc.wb, 0)[0].text, 'bir\niki\n\ndört');
  assert.deepEqual(setShapeText(doc.wb, 0, 0, 'bir\niki\n\ndört'), [], 'the same text');
  const saved = await save(doc);
  assert.equal(sheetShapes((await load(saved)).wb, 0)[0].text, 'bir\niki\n\ndört');
  apply(doc, setShapeText(doc.wb, 0, 0, ''));
  assert.equal(shapeOf(doc.wb, 0).text, undefined);
  // An empty text box keeps the look typed text gets: one paragraph that ends in it.
  const xml = await partText(await save(doc), 'xl/drawings/drawing1.xml');
  assert.match(xml, /<a:p><a:endParaRPr lang="en-US" sz="1100"><a:solidFill><a:srgbClr val="000000"\/><\/a:solidFill><\/a:endParaRPr><\/a:p>/);
  assert.equal(shapeOf((await load(await save(doc))).wb, 0).known, true);
  assert.throws(() => checkedShapeText('a\u0001b'), /cannot hold/);
  assert.throws(() => checkedShapeText('\uD800'), /cannot hold/);
  assert.throws(() => checkedShapeText('x'.repeat(40000)), /longer than/);
  assert.throws(() => checkedShapeText('\n'.repeat(2000)), /more than 2000 lines/);
  assert.equal(checkedShapeText('a\tb'), 'a\tb');
});

// ------------------------------------------------------------------ shapes of the file

test('a shape as Excel writes it is not known; it is moved: only its anchor changes; undone, the file is as it was', async () => {
  const file = await fileWith(twoCell([1, 1], [4, 6], excelShape()));
  const doc = await load(file);
  const shape = shapeOf(doc.wb, 0);
  assert.equal(shape.known, undefined, 'a style and a body of its own: more than the model holds');
  assert.equal(shape.plainText, true);
  assert.deepEqual(shape.textLook!.run, { color: '#FFFFFF', size: 14 }, 'the mark that ends the paragraph, over the colour of the style');
  apply(doc, moveShape(doc.wb, 0, 0, { top: 2, left: 2, bottom: 5, right: 5 }));
  const saved = await save(doc);
  assert.deepEqual(await changedParts(file, saved), ['xl/drawings/drawing1.xml']);
  const before = await partText(file, 'xl/drawings/drawing1.xml');
  const after = await partText(saved, 'xl/drawings/drawing1.xml');
  const body = (xml: string) => xml.slice(xml.indexOf('<xdr:sp>'));
  assert.equal(body(after), body(before), 'everything of the shape keeps its bytes');
  assert.match(after, /<xdr:from><xdr:col>2<\/xdr:col>/);
  if (python) assert.deepEqual(packageProblems({ 'moved.xlsx': saved }), []);
  doc.history.undo();
  assert.ok(sameBytes(await save(doc), file), 'undone: byte for byte');
});

test('a shape of the file is turned: only `a:xfrm rot` changes besides its anchor; one without a transform gets one of its turn alone', async () => {
  const bare = sp({ name: 'Bare', props: solid('00FF00') }).replace('cNvPr id="3"', 'cNvPr id="4"');
  const file = await fileWith(twoCell([1, 1], [4, 6], excelShape()) + twoCell([6, 1], [8, 6], bare));
  const doc = await load(file);
  apply(doc, rotateShape(doc.wb, 0, 0, 20));
  apply(doc, rotateShape(doc.wb, 0, 1, 10));
  const saved = await save(doc);
  const after = await partText(saved, 'xl/drawings/drawing1.xml');
  assert.match(after, /<a:xfrm rot="1200000"><a:off x="0" y="0"\/><a:ext cx="0" cy="0"\/><\/a:xfrm>/);
  assert.match(after, /name="Bare"\/><xdr:cNvSpPr\/><\/xdr:nvSpPr><xdr:spPr><a:xfrm rot="600000"\/><a:prstGeom/);
  assert.match(after, /<a:effectRef idx="0">/, 'the style stays');
  const again = await load(saved);
  assert.equal(shapeOf(again.wb, 0).rotation, 20);
  assert.equal(shapeOf(again.wb, 1).rotation, 10);
  if (python) {
    assert.deepEqual(packageProblems({ 'turned.xlsx': saved }), []);
    assert.deepEqual(readOutside({ 'turned.xlsx': saved })['turned.xlsx'].shapes.map((s) => s.rot), ['1200000', '600000']);
  }
  doc.history.undo();
  doc.history.undo();
  assert.ok(sameBytes(await save(doc), file), 'undone: byte for byte');
});

test('the text of a plain shape of the file is replaced: each line a paragraph in the look of the first run, everything else of the body stays', async () => {
  const file = await fileWith(twoCell([1, 1], [4, 6], excelShape()));
  const doc = await load(file);
  apply(doc, setShapeText(doc.wb, 0, 0, 'Bir\nİki'));
  // The model draws the new text in the look of the first run.
  assert.deepEqual(shapeOf(doc.wb, 0).text!.paragraphs.map((p) => [p.align, p.runs.map((r) => [r.text, r.color, r.size])]), [
    ['center', [['Bir', '#FF0000', 14]]],
    ['center', [['İki', '#FF0000', 14]]],
  ]);
  const saved = await save(doc);
  const after = await partText(saved, 'xl/drawings/drawing1.xml');
  const paragraph = (text: string) => `<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>${text}</a:t></a:r><a:endParaRPr lang="tr-TR" sz="1400"/></a:p>`;
  assert.ok(after.includes(`<a:bodyPr vertOverflow="clip" wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/>${paragraph('Bir')}${paragraph('İki')}</xdr:txBody>`), after);
  const again = await load(saved);
  assert.equal(sheetShapes(again.wb, 0)[0].text, 'Bir\nİki');
  assert.equal(shapeOf(again.wb, 0).plainText, true);
  // Taken away: one paragraph with the mark that ends it.
  apply(doc, setShapeText(doc.wb, 0, 0, ''));
  assert.match(await partText(await save(doc), 'xl/drawings/drawing1.xml'), /<a:lstStyle\/><a:p><a:pPr algn="ctr"\/><a:endParaRPr lang="tr-TR" sz="1400"\/><\/a:p><\/xdr:txBody>/);
  if (python) assert.deepEqual(packageProblems({ 'text.xlsx': saved }), []);
  doc.history.undo();
  doc.history.undo();
  assert.ok(sameBytes(await save(doc), file), 'undone: byte for byte');
});

test('a shape of the file whose text has two looks, a field or the text of a cell keeps its text; its look and copies are refused for any shape of the file the engine does not know', async () => {
  const mixed = txBody(para(run('A', 'b="1"') + run('B')));
  const field = txBody(para('<a:fld id="{1}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld>'));
  const linked = excelShape({ id: 5, name: 'Linked' }).replace('<xdr:sp>', '<xdr:sp macro="" textlink="$A$1">');
  const file = await fileWith(twoCell([1, 1], [2, 2], excelShape({ id: 3, body: mixed })) + twoCell([3, 1], [4, 2], excelShape({ id: 4, body: field })) + twoCell([5, 1], [6, 2], linked));
  const doc = await load(file);
  for (const id of [0, 1, 2]) {
    assert.equal(shapeOf(doc.wb, id).plainText, undefined, `shape ${id}`);
    assert.match(shapeEditProblem(doc.wb, 0, id, 'text') || '', /Cannot change the text of this shape/);
    assert.throws(() => setShapeText(doc.wb, 0, id, 'x'), /Cannot change the text of this shape/);
    assert.match(shapeFormatProblem(doc.wb, 0, id) || '', /Cannot change the look of this shape/);
    assert.throws(() => formatShape(doc.wb, 0, id, { fill: '#000000' }), /Cannot change the look/);
    assert.match(shapeCopyProblem(doc.wb, 0, id) || '', /Cannot copy this shape/);
    assert.throws(() => copyShape(doc.wb, 0, id), /Cannot copy this shape/);
  }
  const info = sheetShapes(doc.wb, 0);
  assert.deepEqual(info.map((shape) => [shape.ofFile, shape.formattable, shape.turnable, shape.textEditable, shape.removable]), [
    [true, false, true, false, true],
    [true, false, true, false, true],
    [true, false, true, false, true],
  ]);
});

test('a shape the file holds in a wrapper of alternatives is moved and deleted, not turned nor given text', async () => {
  const inner = twoCell([1, 1], [4, 6], excelShape());
  const file = await fileWith(`<mc:AlternateContent><mc:Choice Requires="a14">${inner}</mc:Choice><mc:Fallback>${inner.replace('cNvPr id="3"', 'cNvPr id="4"')}</mc:Fallback></mc:AlternateContent>`);
  const doc = await load(file);
  assert.equal(doc.wb.sheets[0].drawings![0].alternative, true);
  assert.match(shapeEditProblem(doc.wb, 0, 0) || '', /another form/);
  assert.throws(() => rotateShape(doc.wb, 0, 0, 10), /another form/);
  assert.throws(() => setShapeText(doc.wb, 0, 0, 'x'), /another form/);
  apply(doc, moveShape(doc.wb, 0, 0, { top: 5, left: 5, bottom: 7, right: 7 }));
  const saved = await save(doc);
  if (python) assert.deepEqual(packageProblems({ 'alternative.xlsx': saved }), []);
  apply(doc, removeShape(doc.wb, 0, 0));
  const gone = await save(doc);
  assert.equal(ZipArchive.read(gone).has('xl/drawings/drawing1.xml'), false, 'nothing is left to draw');
  if (python) assert.deepEqual(packageProblems({ 'gone.xlsx': gone }), []);
});

test('a shape of the file the engine wrote is known: changed in every way, it is written anew with the id it has', async () => {
  const doc = await blank();
  apply(doc, addShape(doc.wb, 0, 'roundRect', undefined, { text: 'Eski' }));
  apply(doc, addShape(doc.wb, 0, 'arrow'));
  const file = await save(doc);
  const again = await load(file);
  assert.equal(shapeOf(again.wb, 0).known, true);
  apply(again, formatShape(again.wb, 0, 0, { fill: '#70AD47', textColor: '#000000' }));
  apply(again, setShapeText(again.wb, 0, 0, 'Yeni'));
  apply(again, formatShape(again.wb, 0, 1, { lineColor: '#FF0000', lineWidth: 2 }));
  const saved = await save(again);
  const xml = await partText(saved, 'xl/drawings/drawing1.xml');
  assert.match(xml, /<xdr:cNvPr id="2" name="Rectangle: Rounded Corners 1"\/>/);
  assert.match(xml, /<a:srgbClr val="70AD47"\/>/);
  assert.match(xml, /<a:t>Yeni<\/a:t>/);
  assert.match(xml, /<a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"\/><\/a:solidFill><a:tailEnd type="triangle"\/><\/a:ln>/);
  const third = await load(saved);
  assert.equal(shapeOf(third.wb, 0).known, true);
  assert.equal(shapeOf(third.wb, 0).fill, '#70AD47');
  assert.equal(sheetShapes(third.wb, 0)[0].text, 'Yeni');
  assert.deepEqual(sheetShapes(third.wb, 0)[1].outline, { color: '#FF0000', width: 2, tail: 'triangle' });
  if (python) assert.deepEqual(packageProblems({ 'known.xlsx': saved }), []);
  again.history.undo();
  again.history.undo();
  again.history.undo();
  assert.ok(sameBytes(await save(again), file), 'undone: byte for byte');
});

test('the reader knows a shape only when its element is just what the writer makes: one attribute more and it is not', () => {
  const shape: DrawingShape = { type: 'shape', geometry: 'rect', fill: '#112233', name: 'Rectangle 1', textLook: { run: { size: 11 } } };
  const spaces = 'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
  const parse = (xml: string) => parseXml(`<w ${spaces}>${xml}</w>`).root.children.find((node) => node.type === 'element') as never;
  const made = shapeElementXml(shape, 2, { x: 0, y: 0, width: 10, height: 10 });
  // The place the transform gives is not compared: the anchor says where a shape is.
  assert.equal(sameShapeElement(parse(made), parse(shapeElementXml(shape, 2, { x: 5, y: 5, width: 50, height: 50 }))), true);
  assert.equal(sameShapeElement(parse(made), parse(made.replace('<a:prstGeom prst="rect">', '<a:prstGeom prst="rect" foo="1">'))), false);
  assert.equal(sameShapeElement(parse(made), parse(made.replace('</xdr:spPr>', '<a:effectLst/></xdr:spPr>'))), false);
  assert.equal(sameShapeElement(parse(made), parse(made.replace('id="2"', 'id="3"'))), false);
});

test('a shape the writer makes of a model without insets is read back as known: the body states the defaults', async () => {
  const doc = await blank();
  const ops = addShape(doc.wb, 0, 'rect', undefined, { text: 'İç boşluksuz' });
  apply(doc, ops);
  // A model that says nothing of the space around its text (one built anew from another format, say).
  const shape = shapeOf(doc.wb, 0);
  delete shape.text!.insets;
  delete shape.textLook!.insets;
  const again = await load(await save(doc));
  assert.equal(shapeOf(again.wb, 0).known, true);
  assert.deepEqual(shapeOf(again.wb, 0).text!.insets, [9.6, 4.8, 9.6, 4.8]);
});

test('reading a drawing of many shapes is quick: what the writer never writes (a style) says at once that a shape is not known', async () => {
  const { parseDrawing } = await import('../src/formats/xlsx/drawings');
  const { drawingXml } = await import('./sheet-drawing-fixtures');
  const own = shapeElementXml({ type: 'shape', geometry: 'rect', fill: '#4472C4', text: { paragraphs: [{ runs: [{ text: 'Metin', size: 11 }] }] }, textLook: { run: { size: 11 } }, name: 'R' }, 3, { x: 0, y: 0, width: 0, height: 0 });
  for (const inner of [own, excelShape()]) {
    const xml = drawingXml(Array.from({ length: 5000 }, (_, i) => twoCell([1, i], [3, i + 2], inner)).join(''));
    const root = parseXml(new TextEncoder().encode(xml)).root;
    const start = Date.now();
    const read = parseDrawing(root, { rels: [], themeColors: [], warnings: [] });
    // A guard against a stall, not a measure of speed: about ten times what it takes on an idle machine.
    assert.ok(Date.now() - start < 3000, `${Date.now() - start} ms`);
    assert.equal(read.length, 5000);
    assert.equal(read.every((drawing) => (drawing.object as DrawingShape).known === (inner === own ? true : undefined)), true);
  }
});

// ------------------------------------------------------------------ copies

test('a shape is copied and pasted: beside where it stood, or at a cell, into any workbook; a cut deletes it as one step', async () => {
  const doc = await blank();
  apply(doc, addShape(doc.wb, 0, 'ellipse', at(50, 50, 120, 60), { text: 'Kopya', fill: '#FFC000' }));
  apply(doc, rotateShape(doc.wb, 0, 0, 15));
  const copy = copyShape(doc.wb, 0, 0);
  const pasted = pasteShape(doc.wb, 0, copy);
  apply(doc, pasted.ops);
  assert.equal(pasted.id, 1);
  assert.deepEqual(boxOf(doc.wb, 1), { x: 50 + SHAPE_COPY_OFFSET, y: 50 + SHAPE_COPY_OFFSET, width: 120, height: 60 });
  assert.equal(shapeOf(doc.wb, 1).rotation, 15);
  assert.equal(shapeOf(doc.wb, 1).fill, '#FFC000');
  assert.equal(shapeOf(doc.wb, 1).name, 'Oval 2');
  // Into another workbook, at a cell.
  const other = await blank();
  apply(other, pasteShape(other.wb, 0, copy, { row: 3, col: 2 }).ops);
  const metrics = createSheetMetrics(other.wb.sheets[0]);
  assert.deepEqual(boxOf(other.wb, 0), { x: metrics.cols.offset(2), y: metrics.rows.offset(3), width: 120, height: 60 });
  assert.equal(sheetShapes(other.wb, 0)[0].text, 'Kopya');
  // A cut: one step that deletes it.
  const cut = cutShape(doc.wb, 0, 0);
  apply(doc, cut.ops);
  assert.equal(doc.wb.sheets[0].drawings!.length, 1);
  // What a copy that comes from anywhere has to be.
  assert.throws(() => pasteShape(doc.wb, 0, { shape: { type: 'picture' } } as unknown as ShapeCopy), /not a shape that was copied/);
  assert.throws(() => pasteShape(doc.wb, 0, { ...copy, shape: { ...copy.shape, fill: 'url(x)' } }), /not a shape that was copied/);
  assert.throws(() => pasteShape(doc.wb, 0, copy, { row: -1, col: 0 }), /a cell of the sheet/);
  const saved = await save(other);
  if (python) assert.deepEqual(packageProblems({ 'pasted.xlsx': saved }), []);
});

// ------------------------------------------------------------------ what refuses

test('a protected sheet (objects), and a room, take no shape', async () => {
  const file = await pictureWorkbook({ protection: '<sheetProtection sheet="1" objects="1"/>' });
  const doc = await load(file);
  assert.equal(shapesProblem(doc.wb, 0), 'Cannot change the shapes of this sheet: it is protected');
  assert.throws(() => addShape(doc.wb, 0, 'rect'), /it is protected/);
  const free = await blank();
  const ops = addShape(free.wb, 0, 'rect');
  assert.throws(() => encodeOp(free.wb, ops[0]), (error: unknown) => error instanceof CollabUnsupportedError && error.what === 'picture' && /Shapes cannot be added/.test(error.message));
  apply(free, ops);
  for (const made of [moveShapeTo(free.wb, 0, 0, { x: 5, y: 5, width: 50, height: 50 }), rotateShape(free.wb, 0, 0, 5), setShapeText(free.wb, 0, 0, 'x'), formatShape(free.wb, 0, 0, { fill: '#000000' }), removeShape(free.wb, 0, 0), pasteShape(free.wb, 0, copyShape(free.wb, 0, 0)).ops]) {
    assert.throws(() => encodeOp(free.wb, made[0]), CollabUnsupportedError);
  }
  assert.throws(() => addShape(free.wb, 0, 'star' as never), /A shape is a rectangle/);
  assert.throws(() => addShape(free.wb, 0, 'rect', undefined, { name: 'x'.repeat(300) }), /longer than 255/);
});

test('a workbook built anew (another format saved as xlsx) writes its shapes with the same writer: read back, they are known', async () => {
  const doc = await blank();
  apply(doc, addShape(doc.wb, 0, 'textBox', undefined, { text: 'Kutu' }));
  apply(doc, addShape(doc.wb, 0, 'doubleArrow'));
  const built = await buildXlsx(doc.wb);
  const again = await load(built);
  assert.equal(shapeOf(again.wb, 0).known, true);
  assert.equal(shapeOf(again.wb, 0).textBox, true);
  assert.equal(shapeOf(again.wb, 1).line!.head, 'triangle');
  if (python) assert.deepEqual(packageProblems({ 'built.xlsx': built }), []);
});

// ------------------------------------------------------------------ drawing

test('a line is found near it, a turned shape where it is drawn; the ends of a line are drawn', async () => {
  const doc = await blank();
  apply(doc, addShape(doc.wb, 0, 'arrow', at(100, 100, 200, 0)));
  apply(doc, addShape(doc.wb, 0, 'rect', at(400, 100, 200, 40)));
  apply(doc, rotateShape(doc.wb, 0, 1, 90));
  const sheet = doc.wb.sheets[0];
  const layer = new SheetDrawings(sheet);
  const metrics = createSheetMetrics(sheet);
  assert.equal(layer.hit(metrics, 150, 102), 0, 'two pixels off a flat line');
  assert.equal(layer.hit(metrics, 150, 110), -1, 'ten pixels off');
  assert.equal(layer.hit(metrics, 500, 30), 1, 'above the box it had: inside the box it is turned into');
  assert.equal(layer.hit(metrics, 410, 120), -1, 'inside the box it had, beside the box it is turned into');
  // The arrow head is painted at the end of the line: a fill is asked for, around (300, 100).
  const calls: string[] = [];
  const ctx = new Proxy({} as Record<string, unknown>, {
    get(target, key: string) {
      if (key in target) return target[key];
      return (...args: unknown[]) => {
        calls.push(`${key}(${args.map((a) => (typeof a === 'number' ? Math.round(a) : '')).join(',')})`);
        return key === 'measureText' ? { width: 0 } : undefined;
      };
    },
    set(target, key: string, value) {
      target[key] = value;
      return true;
    },
  });
  layer.paint(ctx as unknown as CanvasRenderingContext2D, { metrics, clip: { x: 0, y: 0, width: 1000, height: 1000 }, originX: 0, originY: 0, devicePixelRatio: 1 } as never);
  const tip = calls.indexOf('moveTo(300,100)');
  assert.ok(tip >= 0 && calls.slice(tip).some((call) => call.startsWith('fill(')), calls.join(' '));
});
