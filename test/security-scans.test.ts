// Text scans of parts a file brings must take time in proportion to the part, whatever the
// part holds: a pattern that reads "up to the next >" from each of many starts that are
// never closed reads the rest of the part from each of them. Each case is a small file
// built to be that, timed on reading or on writing; the limits are guards against a hang
// (about ten times what an idle machine needs), not measurements of speed.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { numberedName, writeXlsx } from '../src/formats/xlsx/write';
import { open, ZipArchive } from '../src/index';
import type { OfficeFile, Workbook } from '../src/index';
import { setComment } from '../src/ops/notes';
import { History, type Op } from '../src/ops/ops';
import { utf8 } from './fixtures';
import { comment, commentsXml, drawingWorkbook as bareWorkbook, vmlNote, vmlXml, type DrawingBookSpec } from './sheet-drawing-fixtures';

const decoder = new TextDecoder();
const REPEAT = 64000;
/** Ten times what reading or writing one of these files takes on an idle machine. */
const LIMIT_MS = 500;

async function drawingWorkbook(spec: DrawingBookSpec): Promise<Uint8Array> {
  const zip = ZipArchive.read(await bareWorkbook(spec));
  const types = decoder.decode(await zip.read('[Content_Types].xml'));
  zip.set('[Content_Types].xml', utf8(types.replace('</Types>', '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>')));
  return zip.write();
}

interface Loaded {
  file: OfficeFile;
  wb: Workbook;
  history: History;
}

async function load(bytes: Uint8Array): Promise<Loaded> {
  const file = await open(bytes);
  const wb = await file.workbook();
  return { file, wb, history: new History(wb) };
}

async function save(doc: Loaded): Promise<Uint8Array> {
  await writeXlsx(doc.file.package!, doc.wb);
  return doc.file.package!.save();
}

const apply = (doc: Loaded, ops: Op[]) => doc.history.apply(ops, undefined, undefined);

async function timed<T>(work: () => Promise<T>): Promise<{ value: T; ms: number }> {
  const start = Date.now();
  const value = await work();
  return { value, ms: Date.now() - start };
}

/** A box of a note that is not well-formed: many <x:ClientData starts and no ">" until the end of the shape. */
const unclosedClientData = () => `<v:shape id="_x0000_s1100" type="#_x0000_t202">${'<x:ClientData '.repeat(REPEAT)}</v:shape>`;

test('reading: a box of a note with many unclosed <x:ClientData starts is read in linear time', async () => {
  const bytes = await drawingWorkbook({
    sheet: '<sheetData/>',
    comments: commentsXml(['a'], comment('A1', 0, '<t>x</t>')),
    vml: vmlXml(unclosedClientData() + vmlNote(0, 0)),
  });
  const { value: wb, ms } = await timed(async () => (await open(bytes)).workbook());
  assert.ok(ms < LIMIT_MS, `read in ${ms} ms`);
  // The box after it is still the one of the note.
  assert.equal(wb.sheets[0].comments?.length, 1);
});

test('reading: the attribute that makes a box a note is found only in the start tag of <x:ClientData>', async () => {
  // The text after the tag is not one of its attributes.
  const shape = '<v:shape id="_x0000_s1100" type="#_x0000_t202"><x:ClientData ObjectType="Drop"><x:Row>0</x:Row><x:Column>0</x:Column><x:Note ObjectType="Note"/></x:ClientData></v:shape';
  const bytes = await drawingWorkbook({
    sheet: '<sheetData/>',
    comments: commentsXml(['a'], comment('A1', 0, '<t>x</t>')),
    // The shape is not closed, so the part is not well-formed and is read as text.
    vml: vmlXml(`${shape} </v:shape>${vmlNote(0, 0, { visible: true })}`),
  });
  const wb = await (await open(bytes)).workbook();
  assert.equal(wb.sheets[0].comments?.length, 1);
});

test('writing: boxes of notes move with rows in linear time in a legacy drawing that is not well-formed', async () => {
  const bytes = await drawingWorkbook({
    sheet: '<sheetData/>',
    comments: commentsXml(['a'], comment('B2', 0, '<t>x</t>')),
    vml: vmlXml(unclosedClientData() + vmlNote(1, 1)),
  });
  const doc = await load(bytes);
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }]);
  const { value: saved, ms } = await timed(() => save(doc));
  assert.ok(ms < LIMIT_MS, `written in ${ms} ms`);
  const vml = decoder.decode(await ZipArchive.read(saved).read('xl/drawings/vmlDrawing1.vml'));
  assert.ok(vml.includes('<x:Row>2</x:Row>'), 'the box of the note moved down a row');
});

test('writing: the box of a deleted note is taken out in linear time in a legacy drawing that is not well-formed', async () => {
  const bytes = await drawingWorkbook({
    sheet: '<sheetData/>',
    comments: commentsXml(['a'], comment('A1', 0, '<t>x</t>') + comment('C3', 0, '<t>y</t>')),
    vml: vmlXml(unclosedClientData() + vmlNote(0, 0) + vmlNote(2, 2)),
  });
  const doc = await load(bytes);
  apply(doc, setComment(doc.wb, 0, 0, 0, undefined));
  const { value: saved, ms } = await timed(() => save(doc));
  assert.ok(ms < LIMIT_MS, `written in ${ms} ms`);
  const vml = decoder.decode(await ZipArchive.read(saved).read('xl/drawings/vmlDrawing1.vml'));
  assert.ok(!vml.includes('<x:Row>0</x:Row>'), 'the box of A1 is gone');
  assert.ok(vml.includes('<x:Row>2</x:Row>'), 'the box of C3 stays');
});

test('writing: a note is added in linear time to a legacy drawing with many unclosed <o:idmap starts', async () => {
  const idmaps = `<!--${'<o:idmap :idmap '.repeat(REPEAT)}-->`;
  const vml = vmlXml(vmlNote(0, 0)).replace('<o:shapelayout', `${idmaps}<o:shapelayout`);
  const bytes = await drawingWorkbook({ sheet: '<sheetData/>', comments: commentsXml(['a'], comment('A1', 0, '<t>x</t>')), vml });
  const doc = await load(bytes);
  apply(doc, setComment(doc.wb, 0, 4, 4, 'new', { author: 'b' }));
  const { value: saved, ms } = await timed(() => save(doc));
  assert.ok(ms < LIMIT_MS, `written in ${ms} ms`);
  const written = await (await open(saved)).workbook();
  assert.equal(written.sheets[0].comments?.length, 2);
});

test('writing: the <legacyDrawing> of a sheet is found in linear time past comments that repeat its start', async () => {
  const bytes = await drawingWorkbook({
    sheet: `<sheetData/><!--${'<legacyDrawing '.repeat(REPEAT)}-->`,
    comments: commentsXml(['a'], comment('A1', 0, '<t>x</t>')),
    vml: vmlXml(vmlNote(0, 0)),
  });
  const doc = await load(bytes);
  apply(doc, setComment(doc.wb, 0, 3, 3, 'new', { author: 'b' }));
  const { value: saved, ms } = await timed(() => save(doc));
  assert.ok(ms < LIMIT_MS, `written in ${ms} ms`);
  const vml = decoder.decode(await ZipArchive.read(saved).read('xl/drawings/vmlDrawing1.vml'));
  assert.ok(vml.includes('<x:Row>3</x:Row>'), 'the box of the new note is in the legacy drawing');
});

test('writing: a <legacyDrawing> inside a comment of the sheet is not the one of the sheet', async () => {
  const bytes = await drawingWorkbook({
    sheet: '<sheetData/><!-- <legacyDrawing r:id="rIdGone"/> -->',
    comments: commentsXml(['a'], comment('A1', 0, '<t>x</t>')),
    vml: vmlXml(vmlNote(0, 0)),
  });
  const doc = await load(bytes);
  apply(doc, setComment(doc.wb, 0, 3, 3, 'new', { author: 'b' }));
  const saved = await save(doc);
  const zip = ZipArchive.read(saved);
  const vml = decoder.decode(await zip.read('xl/drawings/vmlDrawing1.vml'));
  assert.ok(vml.includes('<x:Row>3</x:Row>'), 'the box of the new note went into the drawing the sheet names');
  const written = await (await open(saved)).workbook();
  assert.equal(written.sheets[0].comments?.length, 2);
});

test('numberedName(): the stem without its number, and the extension', () => {
  assert.deepEqual(numberedName('/xl/charts/colors12.xml'), { stem: '/xl/charts/colors', extension: '.xml' });
  assert.deepEqual(numberedName('/xl/embeddings/Microsoft_Excel_Worksheet3.xlsx'), { stem: '/xl/embeddings/Microsoft_Excel_Worksheet', extension: '.xlsx' });
  assert.deepEqual(numberedName('/xl/charts/style'), { stem: '/xl/charts/style', extension: '' });
  assert.deepEqual(numberedName('/xl/a.b/item7'), { stem: '/xl/a.b/item', extension: '' });
  assert.deepEqual(numberedName('/xl/charts/x.tar.gz'), { stem: '/xl/charts/x.tar', extension: '.gz' });
  // A line break in a name is no obstacle (a pattern with "." stopped at it).
  assert.deepEqual(numberedName('/xl/a\nb5.xml'), { stem: '/xl/a\nb', extension: '.xml' });
});

test('numberedName(): a long name of digits is split in linear time', () => {
  const name = `/xl/${'1'.repeat(200000)}x.xml`;
  const start = Date.now();
  const split = numberedName(name);
  assert.ok(Date.now() - start < 50, `split in ${Date.now() - start} ms`);
  assert.equal(split.extension, '.xml');
  assert.equal(split.stem, name.slice(0, -4));
});
