// Drawings of a sheet: what the reader makes of a drawing part, and that the
// view moves anchors after structural edits exactly as the writer moves them
// in the file.

import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { test } from 'node:test';
import { fileURLToPath } from 'node:url';
import { png } from '../e2e/fixtures/docx';
import { applyOp, open, type Workbook } from '../src/index';
import { MAX_GROUP_DEPTH, MAX_SHEET_DRAWINGS, drawingColor, parseDrawing } from '../src/formats/xlsx/drawings';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import type { DrawingAnchor, DrawingGroup, DrawingPicture, DrawingShape, SheetDrawing, SheetStructuralEdit } from '../src/model/sheet/types';
import { sniffImageType } from '../src/render/images';
import { SheetDrawings, adjustDrawingAnchor, anchorRect } from '../src/render/sheet-drawings';
import { parseXml } from '../src/xml';
import {
  PX,
  absolute,
  chartFrame,
  connector,
  drawingWorkbook,
  drawingXml,
  group,
  oneCell,
  outline,
  para,
  pic,
  run,
  solid,
  sp,
  twoCell,
  txBody,
  xfrm,
} from './sheet-drawing-fixtures';

const RED: [number, number, number] = [220, 0, 0];

async function read(spec: Parameters<typeof drawingWorkbook>[0]): Promise<Workbook> {
  return (await open(await drawingWorkbook(spec))).workbook();
}

const drawingsOf = (workbook: Workbook): SheetDrawing[] => workbook.sheets[0].drawings || [];

test('drawings: a picture in a two cell anchor is read with its media part, name and description', async () => {
  const bytes = png(8, 4, RED);
  const workbook = await read({
    drawing: twoCell([1, 2, 5, 3], [3, 6, 10, 4], pic('rIdImg', { name: 'Logo', descr: 'Company logo' })),
    media: { rIdImg: bytes },
  });
  const drawings = drawingsOf(workbook);
  assert.equal(drawings.length, 1);
  assert.deepEqual(drawings[0].anchor, {
    kind: 'twoCell',
    from: { col: 1, colOff: 5 * PX, row: 2, rowOff: 3 * PX },
    to: { col: 3, colOff: 10 * PX, row: 6, rowOff: 4 * PX },
  });
  const picture = drawings[0].object as DrawingPicture;
  assert.equal(picture.type, 'picture');
  assert.equal(picture.name, 'Logo');
  assert.equal(picture.description, 'Company logo');
  assert.equal(picture.media, '/xl/media/rIdImg.bin');
  assert.equal(picture.external, undefined);
  // The bytes come from the package when asked for, not before.
  assert.ok(workbook.media);
  assert.deepEqual(await workbook.media.readPart('/xl/media/rIdImg.bin'), bytes);
  // Nothing but the pictures of drawings can be asked for through it.
  await assert.rejects(workbook.media.readPart('/xl/workbook.xml'));
  assert.equal(workbook.sheets[0].warnings, undefined);
});

test('drawings: one cell and absolute anchors', async () => {
  const workbook = await read({
    drawing: oneCell([2, 3, 4, 6], 100, 50, pic('rIdImg')) + absolute(30, 40, 120, 60, sp({ props: solid('FF0000') })),
    media: { rIdImg: png(2, 2, RED) },
  });
  const drawings = drawingsOf(workbook);
  assert.deepEqual(drawings[0].anchor, { kind: 'oneCell', from: { col: 2, colOff: 4 * PX, row: 3, rowOff: 6 * PX }, width: 100 * PX, height: 50 * PX });
  assert.deepEqual(drawings[1].anchor, { kind: 'absolute', x: 30 * PX, y: 40 * PX, width: 120 * PX, height: 60 * PX });
  const metrics = createSheetMetrics(workbook.sheets[0]);
  assert.deepEqual(anchorRect(drawings[0].anchor, metrics), { x: 2 * 64 + 4, y: 3 * 20 + 6, width: 100, height: 50 });
  assert.deepEqual(anchorRect(drawings[1].anchor, metrics), { x: 30, y: 40, width: 120, height: 60 });
});

test('drawings: editAs is kept; an unknown value reads as the default', async () => {
  const workbook = await read({
    drawing:
      twoCell([0, 0], [1, 1], sp(), 'oneCell') + twoCell([0, 0], [1, 1], sp(), 'absolute') + twoCell([0, 0], [1, 1], sp(), 'twoCell') + twoCell([0, 0], [1, 1], sp(), 'bogus'),
  });
  assert.deepEqual(
    drawingsOf(workbook).map((drawing) => drawing.anchor.editAs),
    ['oneCell', 'absolute', undefined, undefined],
  );
});

test('drawings: crop, rotation and flips of a picture', async () => {
  const workbook = await read({
    drawing: twoCell([0, 0], [2, 4], pic('rIdImg', { fill: '<a:srcRect l="10000" t="20000" r="30000" b="5000"/><a:stretch><a:fillRect/></a:stretch>', xfrm: 'rot="5400000" flipH="1"' })),
    media: { rIdImg: png(2, 2, RED) },
  });
  const picture = drawingsOf(workbook)[0].object as DrawingPicture;
  assert.deepEqual(picture.crop, { left: 0.1, top: 0.2, right: 0.3, bottom: 0.05 });
  assert.equal(picture.rotation, 90);
  assert.equal(picture.flipH, true);
  assert.equal(picture.flipV, undefined);
});

test('drawings: a crop that leaves nothing, or is negative, is not applied', async () => {
  const workbook = await read({
    drawing:
      twoCell([0, 0], [2, 4], pic('rIdImg', { fill: '<a:srcRect l="60000" r="60000"/>' })) + twoCell([0, 0], [2, 4], pic('rIdImg', { fill: '<a:srcRect l="-5000" t="-5000"/>' })),
    media: { rIdImg: png(2, 2, RED) },
  });
  assert.equal((drawingsOf(workbook)[0].object as DrawingPicture).crop, undefined);
  assert.equal((drawingsOf(workbook)[1].object as DrawingPicture).crop, undefined);
});

test('drawings: a linked picture is marked external and has no media; so is an embed that points outside', async () => {
  const workbook = await read({
    drawing: twoCell([0, 0], [2, 4], pic('rIdLink', { link: true })) + twoCell([0, 0], [2, 4], pic('rIdOut')) + twoCell([0, 0], [2, 4], pic('rIdMissing')),
    drawingRels:
      '<Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.com/a.png" TargetMode="External"/>' +
      '<Relationship Id="rIdOut" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="file:///c:/a.png" TargetMode="External"/>' +
      '<Relationship Id="rIdMissing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/gone.png"/>',
  });
  const [linked, outside, missing] = drawingsOf(workbook).map((drawing) => drawing.object as DrawingPicture);
  assert.equal(linked.external, true);
  assert.equal(linked.media, undefined);
  assert.equal(outside.external, true);
  assert.equal(outside.media, undefined);
  // A part that is not in the package: a placeholder, not an error.
  assert.equal(missing.media, undefined);
  assert.equal(missing.external, undefined);
  assert.equal(workbook.media, undefined);
});

test('drawings: shape fill, outline and geometry', async () => {
  const workbook = await read({
    drawing:
      twoCell([0, 0], [2, 4], sp({ prst: 'ellipse', props: solid('00B050') + outline('112233', 2) })) +
      twoCell([0, 0], [2, 4], sp({ prst: 'roundRect', props: '<a:noFill/>' + outline('FF0000', 1, '<a:prstDash val="dash"/>') })) +
      twoCell([0, 0], [2, 4], sp({ props: '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln><a:noFill/></a:ln>' })) +
      twoCell([0, 0], [2, 4], sp({ props: '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="ABCDEF"/></a:gs><a:gs pos="100000"><a:srgbClr val="000000"/></a:gs></a:gsLst></a:gradFill>' })) +
      twoCell([0, 0], [2, 4], sp({ prst: 'not a name!' })),
  });
  const shapes = drawingsOf(workbook).map((drawing) => drawing.object as DrawingShape);
  assert.equal(shapes[0].geometry, 'ellipse');
  assert.equal(shapes[0].fill, '#00B050');
  assert.deepEqual(shapes[0].line, { color: '#112233', width: 2 });
  assert.equal(shapes[1].geometry, 'roundRect');
  assert.equal(shapes[1].fill, undefined);
  assert.deepEqual(shapes[1].line, { color: '#FF0000', width: 1, dash: true });
  assert.equal(shapes[2].fill, '#4472C4');
  assert.equal(shapes[2].line, undefined);
  // A gradient shows as its first colour.
  assert.equal(shapes[3].fill, '#ABCDEF');
  assert.equal(shapes[4].geometry, 'rect');
});

test('drawings: the shape style gives fill, outline and text colour where the shape states none', async () => {
  const style =
    '<xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
    '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style>';
  const workbook = await read({
    drawing:
      twoCell([0, 0], [2, 4], sp({ style, body: txBody(para(run('Merhaba'))) })) +
      // What the shape itself says wins over its style.
      twoCell([0, 0], [2, 4], sp({ style, props: '<a:noFill/><a:ln><a:noFill/></a:ln>', body: txBody(para(run('x', '', solid('FF0000')))) })) +
      // idx 0 means "nothing from the theme".
      twoCell([0, 0], [2, 4], sp({ style: '<xdr:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef></xdr:style>' })),
  });
  const shapes = drawingsOf(workbook).map((drawing) => drawing.object as DrawingShape);
  assert.equal(shapes[0].fill, '#4472C4');
  assert.deepEqual(shapes[0].line, { color: '#223962', width: 1.5 });
  assert.equal(shapes[0].text?.paragraphs[0].runs[0].color, '#FFFFFF');
  assert.equal(shapes[1].fill, undefined);
  assert.equal(shapes[1].line, undefined);
  assert.equal(shapes[1].text?.paragraphs[0].runs[0].color, '#FF0000');
  assert.equal(shapes[2].fill, undefined);
  assert.equal(shapes[2].line, undefined);
});

test('drawings: text of a shape with runs, breaks, fields, alignment, anchor and insets', async () => {
  const body = txBody(
    para(run('Kalın', 'sz="1400" b="1"') + run(' eğik', 'i="1" u="sng" strike="sngStrike"', solid('00FF00') + '<a:latin typeface="Arial"/>') + '<a:br/>' + run('yeni')) +
      para('<a:fld id="{1}" type="TxLink"><a:rPr lang="en-US"/><a:t>alan</a:t></a:fld>', '<a:pPr algn="ctr"/>') +
      para('', '<a:pPr algn="r"/>') +
      para(run('tema', '', '<a:latin typeface="+mn-lt"/>'), '<a:pPr algn="l"/>'),
    '<a:bodyPr wrap="none" lIns="0" tIns="19050" rIns="9525" anchor="ctr"/>',
  );
  const workbook = await read({ drawing: twoCell([0, 0], [2, 4], sp({ body })) });
  const text = (drawingsOf(workbook)[0].object as DrawingShape).text;
  assert.ok(text);
  assert.equal(text.anchor, 'middle');
  assert.equal(text.wrap, false);
  assert.deepEqual(text.insets, [0, 2, 1, 4.8]);
  assert.deepEqual(text.paragraphs[0].runs, [
    { text: 'Kalın', size: 14, bold: true },
    { text: ' eğik', italic: true, underline: true, strike: true, color: '#00FF00', font: 'Arial' },
    { text: '\n' },
    { text: 'yeni' },
  ]);
  assert.equal(text.paragraphs[0].align, undefined);
  assert.deepEqual(text.paragraphs[1], { runs: [{ text: 'alan' }], align: 'center' });
  assert.deepEqual(text.paragraphs[2], { runs: [], align: 'right' });
  // A font of the theme is not a font name.
  assert.deepEqual(text.paragraphs[3], { runs: [{ text: 'tema' }] });
});

test('drawings: a shape whose text is empty has no text', async () => {
  const workbook = await read({ drawing: twoCell([0, 0], [2, 4], sp({ body: txBody(para('') + para('<a:br/>')) })) });
  assert.equal((drawingsOf(workbook)[0].object as DrawingShape).text, undefined);
});

test('drawings: a connector is a line shape; a chart frame is a placeholder', async () => {
  const workbook = await read({
    drawing:
      twoCell([0, 0], [2, 4], connector(outline('0000FF', 2), xfrm(0, 0, 10, 10, 'flipV="1"'))) +
      twoCell([0, 0], [2, 4], sp({ prst: 'line', props: outline('000000') })) +
      twoCell([3, 0], [8, 10], chartFrame()) +
      twoCell([3, 0], [8, 10], chartFrame('http://schemas.openxmlformats.org/drawingml/2006/diagram')) +
      twoCell([3, 0], [8, 10], chartFrame('http://schemas.microsoft.com/office/drawing/2010/slicer')),
  });
  const objects = drawingsOf(workbook).map((drawing) => drawing.object);
  assert.deepEqual(objects[0], { name: 'Connector 1', flipV: true, type: 'shape', geometry: 'line', connector: true, line: { color: '#0000FF', width: 2 } });
  assert.equal((objects[1] as DrawingShape).connector, true);
  assert.deepEqual(objects[2], { name: 'Chart 1', type: 'placeholder', what: 'chart' });
  assert.equal(objects[3].type === 'placeholder' && objects[3].what, 'diagram');
  assert.equal(objects[4].type === 'placeholder' && objects[4].what, 'object');
});

test('drawings: members of a group are placed by the child space of the group', async () => {
  const members =
    sp({ xfrm: xfrm(1000, 2000, 500, 250), props: solid('FF0000'), name: 'a' }) +
    sp({ xfrm: xfrm(1500, 2500, 500, 500, 'rot="2700000"'), props: solid('0000FF'), name: 'b' }) +
    // No transform of its own: it has no place in the group.
    sp({ props: solid('00FF00'), name: 'c' }) +
    group(xfrm(1000, 2000, 1000, 1000, '', [0, 0, 100, 100]), sp({ xfrm: xfrm(50, 50, 50, 50), name: 'inner' }));
  const workbook = await read({ drawing: twoCell([0, 0], [4, 8], group(xfrm(0, 0, 9999, 9999, 'flipH="1"', [1000, 2000, 1000, 1000]), members)) });
  const top = drawingsOf(workbook)[0].object as DrawingGroup;
  assert.equal(top.type, 'group');
  assert.equal(top.flipH, true);
  assert.deepEqual(
    top.children.map((member) => [member.name, member.box, member.rotation]),
    [
      ['a', { x: 0, y: 0, width: 0.5, height: 0.25 }, undefined],
      ['b', { x: 0.5, y: 0.5, width: 0.5, height: 0.5 }, 45],
      ['Group 1', { x: 0, y: 0, width: 1, height: 1 }, undefined],
    ],
  );
  const inner = top.children[2] as DrawingGroup;
  assert.deepEqual(inner.children[0].box, { x: 0.5, y: 0.5, width: 0.5, height: 0.5 });
});

test('drawings: groups nested deeper than the limit lose their deepest members, nothing else', () => {
  let inner = sp({ xfrm: xfrm(0, 0, 10, 10), name: 'deepest' });
  for (let i = 0; i < MAX_GROUP_DEPTH + 8; i++) inner = group(xfrm(0, 0, 10, 10, '', [0, 0, 10, 10]), inner);
  const warnings: string[] = [];
  const drawings = parseDrawing(parseXml(drawingXml(twoCell([0, 0], [1, 1], inner) + twoCell([0, 0], [1, 1], sp({ name: 'after' })))).root, { rels: [], themeColors: [], warnings });
  assert.equal(drawings.length, 2);
  let depth = 0;
  let object = drawings[0].object;
  while (object.type === 'group' && object.children.length > 0) {
    object = object.children[0];
    depth++;
  }
  assert.equal(depth, MAX_GROUP_DEPTH);
  assert.equal(object.type, 'group');
  assert.equal(drawings[1].object.name, 'after');
});

test('drawings: of an alternate content one branch is read, never both', async () => {
  const mc = (choice: string, fallback: string) =>
    `<mc:AlternateContent><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
  const workbook = await read({
    drawing:
      // Around the content of an anchor.
      twoCell([0, 0], [2, 4], mc(sp({ name: 'choice' }), sp({ name: 'fallback' }))) +
      // Around whole anchors.
      mc(twoCell([1, 1], [2, 2], sp({ name: 'new' })), twoCell([1, 1], [2, 2], sp({ name: 'old' }))) +
      // A choice that holds nothing known: the fallback is read.
      twoCell([0, 0], [2, 4], mc('<xdr:unknown/>', sp({ name: 'only fallback' }))),
  });
  assert.deepEqual(
    drawingsOf(workbook).map((drawing) => drawing.object.name),
    ['choice', 'new', 'only fallback'],
  );
});

test('drawings: an anchor that cannot be read is left out with a warning; the others are shown', async () => {
  const broken = '<xdr:twoCellAnchor><xdr:from><xdr:col>x</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' + `${sp()}<xdr:clientData/></xdr:twoCellAnchor>`;
  const outside = twoCell([20000, 0], [20001, 4], sp());
  const empty = '<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:clientData/></xdr:twoCellAnchor>';
  const workbook = await read({ drawing: broken + twoCell([0, 0], [2, 4], sp({ name: 'good' })) + outside + empty + absolute(0, 0, -5, 10, sp()) });
  assert.deepEqual(
    drawingsOf(workbook).map((drawing) => drawing.object.name),
    ['good'],
  );
  assert.deepEqual(workbook.sheets[0].warnings, ['Drawing: 4 object(s) could not be read and are not shown']);
});

test('drawings: a drawing part that is not XML leaves the sheet readable, with a warning', async () => {
  const bytes = await drawingWorkbook({ sheet: '<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData>', drawing: '' });
  const file = await open(bytes);
  file.package?.writePart('/xl/drawings/drawing1.xml', new TextEncoder().encode('<xdr:wsDr><broken'));
  const reread = await open(await file.save());
  const workbook = await reread.workbook();
  assert.equal(workbook.sheets[0].drawings, undefined);
  assert.equal(workbook.sheets[0].rows.get(0)?.cells[0].value, 7);
  assert.equal(workbook.sheets[0].warnings?.length, 1);
  assert.match(workbook.sheets[0].warnings?.[0] ?? '', /^Drawing could not be read: /);
});

test('drawings: no more anchors than the limit are read', () => {
  const one = twoCell([0, 0], [1, 1], sp());
  const warnings: string[] = [];
  const drawings = parseDrawing(parseXml(drawingXml(one.repeat(MAX_SHEET_DRAWINGS + 5))).root, { rels: [], themeColors: [], warnings });
  assert.equal(drawings.length, MAX_SHEET_DRAWINGS);
  assert.equal(warnings.length, 1);
});

test('drawings: offsets are clamped, never negative or NaN', async () => {
  const anchor =
    '<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>-9525</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>abc</xdr:rowOff></xdr:from>' +
    '<xdr:to><xdr:col>2</xdr:col><xdr:colOff>99999999999999</xdr:colOff><xdr:row>2</xdr:row></xdr:to>' +
    `${sp()}<xdr:clientData/></xdr:twoCellAnchor>`;
  const workbook = await read({ drawing: anchor });
  const read1 = drawingsOf(workbook)[0].anchor;
  assert.deepEqual(read1.from, { col: 1, colOff: 0, row: 1, rowOff: 0 });
  assert.equal(read1.to?.rowOff, 0);
  // However far the offset claims to reach, the corner stays inside its cell.
  const rect = anchorRect(read1, createSheetMetrics(workbook.sheets[0]));
  assert.deepEqual(rect, { x: 64, y: 20, width: 128, height: 20 });
});

test('drawings: colours of the theme with their changes', () => {
  const theme = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
  const color = (xml: string) => drawingColor(theme, parseXml(`<x xmlns:a="a">${xml}</x>`).root.children[0] as never);
  assert.equal(color('<a:srgbClr val="1a2B3c"/>'), '#1A2B3C');
  assert.equal(color('<a:schemeClr val="tx1"/>'), '#000000');
  assert.equal(color('<a:schemeClr val="bg1"/>'), '#FFFFFF');
  assert.equal(color('<a:schemeClr val="accent2"/>'), '#ED7D31');
  assert.equal(color('<a:schemeClr val="bg1"><a:lumMod val="85000"/></a:schemeClr>'), '#D9D9D9');
  assert.equal(color('<a:sysClr val="windowText" lastClr="101010"/>'), '#101010');
  assert.equal(color('<a:prstClr val="red"/>'), '#FF0000');
  assert.equal(color('<a:srgbClr val="FF0000"><a:alpha val="0"/></a:srgbClr>'), undefined);
  assert.equal(color('<a:srgbClr val="javascript:alert(1)"/>'), undefined);
  assert.equal(color('<a:schemeClr val="constructor"/>'), undefined);
  assert.equal(color('<a:prstClr val="toString"/>'), undefined);
  // Without a theme the two colours every file relies on are still known.
  assert.equal(drawingColor([], parseXml('<a:schemeClr xmlns:a="a" val="tx1"/>').root), '#000000');
  assert.equal(drawingColor([], parseXml('<a:schemeClr xmlns:a="a" val="accent1"/>').root), undefined);
});

// --- structural edits: the view moves anchors as the writer does -------------------------------

const ANCHORS: [name: string, xml: string][] = [
  ['twoCell', twoCell([1, 2, 5, 3], [3, 6, 10, 4], sp({ name: 'twoCell' }))],
  ['oneCellEdit', twoCell([1, 2, 5, 3], [3, 6, 10, 4], sp({ name: 'oneCellEdit' }), 'oneCell')],
  ['absoluteEdit', twoCell([1, 2, 5, 3], [3, 6, 10, 4], sp({ name: 'absoluteEdit' }), 'absolute')],
  ['oneCell', oneCell([2, 4, 7, 9], 100, 50, sp({ name: 'oneCell' }))],
  ['absolute', absolute(30, 40, 120, 60, sp({ name: 'absolute' }))],
  ['low', twoCell([0, 9, 1, 1], [1, 12, 2, 2], sp({ name: 'low' }))],
];

const EDIT_CASES: [name: string, edits: SheetStructuralEdit[]][] = [
  ['rows inserted above', [{ kind: 'insert', axis: 'row', at: 0, count: 3 }]],
  ['rows inserted inside', [{ kind: 'insert', axis: 'row', at: 4, count: 2 }]],
  ['rows inserted at the first row of the object', [{ kind: 'insert', axis: 'row', at: 2, count: 1 }]],
  ['rows inserted below', [{ kind: 'insert', axis: 'row', at: 30, count: 2 }]],
  ['columns inserted', [{ kind: 'insert', axis: 'col', at: 2, count: 2 }]],
  ['rows deleted above', [{ kind: 'delete', axis: 'row', at: 0, count: 2 }]],
  ['rows deleted across the top edge', [{ kind: 'delete', axis: 'row', at: 1, count: 3 }]],
  ['rows deleted inside', [{ kind: 'delete', axis: 'row', at: 3, count: 2 }]],
  ['rows deleted across the bottom edge', [{ kind: 'delete', axis: 'row', at: 5, count: 4 }]],
  ['every row of the object deleted', [{ kind: 'delete', axis: 'row', at: 1, count: 8 }]],
  ['columns deleted', [{ kind: 'delete', axis: 'col', at: 1, count: 2 }]],
  [
    'several edits in a row',
    [
      { kind: 'insert', axis: 'row', at: 3, count: 4 },
      { kind: 'delete', axis: 'col', at: 0, count: 1 },
      { kind: 'delete', axis: 'row', at: 2, count: 2 },
      { kind: 'insert', axis: 'col', at: 1, count: 3 },
    ],
  ],
];

for (const [name, edits] of EDIT_CASES) {
  test(`drawings: after ${name} the view shows every anchor where the saved file has it`, async () => {
    const file = await open(await drawingWorkbook({ sheet: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>', drawing: ANCHORS.map(([, xml]) => xml).join('') }));
    const workbook = await file.workbook();
    const sheet = workbook.sheets[0];
    const before = (sheet.drawings || []).map((drawing) => drawing.anchor);
    for (const edit of edits) applyOp(workbook, { type: edit.kind, sheet: 0, axis: edit.axis, at: edit.at, count: edit.count });
    assert.deepEqual(sheet.structuralEdits, edits);
    // The model keeps what the file had; the view applies the edits.
    assert.deepEqual(
      (sheet.drawings || []).map((drawing) => drawing.anchor),
      before,
    );
    const shown = new SheetDrawings(sheet).currentAnchors();

    const reread = await (await open(await file.save())).workbook();
    const saved = (reread.sheets[0].drawings || []).map((drawing) => drawing.anchor);
    assert.deepEqual(
      (reread.sheets[0].drawings || []).map((drawing) => drawing.object.name),
      ANCHORS.map(([anchor]) => anchor),
    );
    assert.deepEqual(shown, saved);

    // And in pixels: the same rectangles on the same rows and columns.
    const metrics = createSheetMetrics(reread.sheets[0]);
    assert.deepEqual(
      new SheetDrawings(sheet).rects(createSheetMetrics(sheet)).map(({ x, y, width, height }) => ({ x, y, width, height })),
      new SheetDrawings(reread.sheets[0]).rects(metrics).map(({ x, y, width, height }) => ({ x, y, width, height })),
    );
  });
}

test('drawings: the rules of a structural edit, spelled out', () => {
  const two: DrawingAnchor = { kind: 'twoCell', from: { col: 1, colOff: 5, row: 2, rowOff: 3 }, to: { col: 3, colOff: 10, row: 6, rowOff: 4 } };
  const insert: SheetStructuralEdit = { kind: 'insert', axis: 'row', at: 4, count: 2 };
  // Moves and sizes with cells: rows inserted inside make it taller.
  assert.deepEqual(adjustDrawingAnchor(two, [insert]).to, { col: 3, colOff: 10, row: 8, rowOff: 4 });
  assert.deepEqual(adjustDrawingAnchor(two, [insert]).from, two.from);
  // Moves but does not size: rows inserted inside change nothing, rows inserted above move both corners.
  const one: DrawingAnchor = { ...two, editAs: 'oneCell' };
  assert.equal(adjustDrawingAnchor(one, [insert]), one);
  const above: SheetStructuralEdit = { kind: 'insert', axis: 'row', at: 0, count: 5 };
  assert.deepEqual(adjustDrawingAnchor(one, [above]), { ...one, from: { ...one.from, row: 7 }, to: { ...one.to, row: 11 } });
  // Does neither.
  const fixed: DrawingAnchor = { ...two, editAs: 'absolute' };
  assert.equal(adjustDrawingAnchor(fixed, [above]), fixed);
  const absoluteAnchor: DrawingAnchor = { kind: 'absolute', x: 1, y: 2, width: 3, height: 4 };
  assert.equal(adjustDrawingAnchor(absoluteAnchor, [above]), absoluteAnchor);
  // A corner whose row is deleted goes to where the deletion was, at the top of that row.
  const cut: SheetStructuralEdit = { kind: 'delete', axis: 'row', at: 1, count: 3 };
  assert.deepEqual(adjustDrawingAnchor(two, [cut]), { kind: 'twoCell', from: { col: 1, colOff: 5, row: 1, rowOff: 0 }, to: { col: 3, colOff: 10, row: 3, rowOff: 4 } });
  // All of its rows deleted: no height left there.
  const all: SheetStructuralEdit = { kind: 'delete', axis: 'row', at: 2, count: 5 };
  const flat = adjustDrawingAnchor(two, [all]);
  assert.deepEqual([flat.from?.row, flat.from?.rowOff, flat.to?.row, flat.to?.rowOff], [2, 0, 2, 0]);
  // Nothing to do: the same object.
  assert.equal(adjustDrawingAnchor(two, []), two);
  assert.equal(adjustDrawingAnchor(two, undefined), two);
  assert.equal(adjustDrawingAnchor(two, [{ kind: 'insert', axis: 'row', at: 100, count: 1 }]), two);
});

test('drawings: undo puts the anchors back', async () => {
  const file = await open(await drawingWorkbook({ drawing: ANCHORS[0][1] }));
  const workbook = await file.workbook();
  const sheet = workbook.sheets[0];
  const layer = new SheetDrawings(sheet);
  const before = layer.currentAnchors()[0];
  const inverse = applyOp(workbook, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 4 });
  assert.equal(layer.currentAnchors()[0].from?.row, 6);
  applyOp(workbook, inverse);
  assert.deepEqual(layer.currentAnchors()[0], before);
});

test('drawings: hidden rows and columns squeeze an object that sizes with cells', async () => {
  const workbook = await read({
    sheet: '<cols><col min="3" max="3" width="9.140625" hidden="1"/></cols><sheetData><row r="4" hidden="1"/><row r="5" hidden="1"/></sheetData>',
    drawing: twoCell([1, 2, 5, 3], [3, 6, 10, 4], sp()) + twoCell([2, 3, 10, 5], [2, 4, 20, 15], sp({ name: 'inside hidden' })),
  });
  const sheet = workbook.sheets[0];
  const rects = new SheetDrawings(sheet).rects(createSheetMetrics(sheet));
  // Column C (64 px) and rows 4 and 5 (20 px each) are gone from between the corners.
  assert.equal(rects.length, 1);
  assert.deepEqual({ x: rects[0].x, y: rects[0].y, width: rects[0].width, height: rects[0].height }, { x: 69, y: 43, width: 128 - 64 + 5, height: 80 - 40 + 1 });
});

test('drawings: a rotation near a quarter turn swaps the sides of the box around its centre', async () => {
  const workbook = await read({ drawing: twoCell([1, 1], [2, 5], sp({ xfrm: xfrm(0, 0, 1, 1, 'rot="5400000"') })) + twoCell([1, 1], [2, 5], sp({ xfrm: xfrm(0, 0, 1, 1, 'rot="1800000"') })) });
  const sheet = workbook.sheets[0];
  const rects = new SheetDrawings(sheet).rects(createSheetMetrics(sheet));
  // The anchor is 64 x 80 at (64, 20): the shape that is turned is 80 x 64 around the same centre.
  assert.deepEqual({ x: rects[0].x, y: rects[0].y, width: rects[0].width, height: rects[0].height }, { x: 56, y: 28, width: 80, height: 64 });
  assert.deepEqual({ x: rects[1].x, y: rects[1].y, width: rects[1].width, height: rects[1].height }, { x: 64, y: 20, width: 64, height: 80 });
});

// --- openpyxl ------------------------------------------------------------------------------

const here = dirname(fileURLToPath(import.meta.url));
const scratch = join(here, '..', 'tmp-drawings-openpyxl');

/** openpyxl needs Pillow to place a picture; without either the cross-check is skipped. */
const hasOpenpyxl = spawnSync('python3', ['-c', 'import openpyxl, PIL'], { encoding: 'utf8' }).status === 0;

const WRITE_SCRIPT = `
import sys, json
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker, AbsoluteAnchor
from openpyxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D
wb = Workbook()
ws = wb.active
ws['A1'] = 1
first = Image(sys.argv[2])
ws.add_image(first, 'C3')
second = Image(sys.argv[2])
second.width, second.height = 120, 45
ws.add_image(second, 'F10')
third = Image(sys.argv[2])
third.anchor = TwoCellAnchor(editAs='oneCell', _from=AnchorMarker(col=1, colOff=19050, row=20, rowOff=28575), to=AnchorMarker(col=4, colOff=95250, row=26, rowOff=0))
ws.add_image(third)
fourth = Image(sys.argv[2])
fourth.anchor = AbsoluteAnchor(pos=XDRPoint2D(x=952500, y=1905000), ext=XDRPositiveSize2D(cx=381000, cy=190500))
ws.add_image(fourth)
other = wb.create_sheet('Diger')
fifth = Image(sys.argv[2])
other.add_image(fifth, 'A1')
wb.save(sys.argv[1])
`;

const READ_SCRIPT = `
import sys, json
from openpyxl import load_workbook
wb = load_workbook(sys.argv[1])
out = []
for ws in wb.worksheets:
    images = []
    for image in ws._images:
        a = image.anchor
        kind = type(a).__name__
        item = {'kind': kind, 'format': image.format, 'size': len(image._data())}
        if hasattr(a, '_from') and a._from is not None:
            item['from'] = [a._from.col, a._from.colOff, a._from.row, a._from.rowOff]
        if kind == 'TwoCellAnchor':
            item['to'] = [a.to.col, a.to.colOff, a.to.row, a.to.rowOff]
            item['editAs'] = a.editAs
        if kind != 'TwoCellAnchor':
            item['ext'] = [a.ext.width, a.ext.height]
        if kind == 'AbsoluteAnchor':
            item['pos'] = [a.pos.x, a.pos.y]
        images.append(item)
    out.append(images)
print(json.dumps(out))
`;

test('drawings: pictures openpyxl writes are read where openpyxl reads them', { skip: hasOpenpyxl ? false : 'python3 with openpyxl and Pillow is not installed' }, async () => {
  mkdirSync(scratch, { recursive: true });
  try {
    const picturePath = join(scratch, 'picture.png');
    const bytes = png(60, 30, RED, [0, 0, 220]);
    writeFileSync(picturePath, bytes);
    const path = join(scratch, 'pictures.xlsx');
    const written = spawnSync('python3', ['-c', WRITE_SCRIPT, path, picturePath], { encoding: 'utf8' });
    assert.equal(written.status, 0, written.stderr);
    const read = spawnSync('python3', ['-c', READ_SCRIPT, path], { encoding: 'utf8' });
    assert.equal(read.status, 0, read.stderr);
    type Expected = { kind: string; format: string; size: number; from?: number[]; to?: number[]; ext?: number[]; pos?: number[]; editAs?: string };
    const expected = JSON.parse(read.stdout) as Expected[][];
    assert.deepEqual(
      expected.map((sheet) => sheet.length),
      [4, 1],
    );
    const workbook = await (await open(new Uint8Array(readFileSync(path)))).workbook();
    // openpyxl groups the anchors by kind when it reads them; what is compared is the set, not the order.
    const sorted = (list: unknown[]) => list.map((item) => JSON.stringify(item)).sort();
    for (let s = 0; s < expected.length; s++) {
      const drawings = workbook.sheets[s].drawings || [];
      const ours: unknown[] = [];
      for (const drawing of drawings) {
        const anchor = drawing.anchor;
        const picture = drawing.object as DrawingPicture;
        assert.equal(picture.type, 'picture');
        // The same picture, byte for byte.
        assert.ok(picture.media);
        const data = await workbook.media!.readPart(picture.media);
        assert.equal(sniffImageType(data), 'image/png');
        const item: Record<string, unknown> = { kind: anchor.kind, size: data.length };
        if (anchor.from) item.from = [anchor.from.col, anchor.from.colOff, anchor.from.row, anchor.from.rowOff];
        if (anchor.to) item.to = [anchor.to.col, anchor.to.colOff, anchor.to.row, anchor.to.rowOff];
        if (anchor.kind === 'twoCell') item.editAs = anchor.editAs || 'twoCell';
        else item.ext = [anchor.width, anchor.height];
        if (anchor.kind === 'absolute') item.pos = [anchor.x, anchor.y];
        ours.push(item);
      }
      const theirs = expected[s].map((want) => {
        const item: Record<string, unknown> = { kind: want.kind === 'TwoCellAnchor' ? 'twoCell' : want.kind === 'OneCellAnchor' ? 'oneCell' : 'absolute', size: want.size };
        if (want.from) item.from = want.from;
        if (want.to) item.to = want.to;
        if (want.kind === 'TwoCellAnchor') item.editAs = want.editAs || 'twoCell';
        else item.ext = want.ext;
        if (want.pos) item.pos = want.pos;
        return item;
      });
      assert.deepEqual(sorted(ours), sorted(theirs));
      assert.equal(workbook.sheets[s].warnings, undefined);
    }
    // In pixels: C3 is at (128, 40) and the first picture keeps its 60 x 30; F10 with a stated size; the absolute one.
    const rects = new SheetDrawings(workbook.sheets[0]).rects(createSheetMetrics(workbook.sheets[0])).map(({ x, y, width, height }) => JSON.stringify({ x, y, width, height }));
    for (const rect of [{ x: 128, y: 40, width: 60, height: 30 }, { x: 5 * 64, y: 9 * 20, width: 120, height: 45 }, { x: 100, y: 200, width: 40, height: 20 }]) {
      assert.ok(rects.indexOf(JSON.stringify(rect)) >= 0, `${JSON.stringify(rect)} among ${rects.join(' ')}`);
    }
  } finally {
    rmSync(scratch, { recursive: true, force: true });
  }
});
