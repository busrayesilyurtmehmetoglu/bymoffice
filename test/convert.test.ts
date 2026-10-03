// convertForEditing(): doc and RTF written out as docx, xls as xlsx, with a
// report of what the new file does not carry over.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { convertForEditing } from '../src/formats/convert';
import { open } from '../src/index';
import type { Paragraph } from '../src/model/doc/types';
import { simpleDoc } from './doc-binary-fixtures';
import { cat, number, workbookStream, xlsFile } from './xls-fixtures';

const rtf = (body: string) => new TextEncoder().encode(`{\\rtf1\\ansi ${body}\\par}`);
const SHAPE = '{\\shp{\\*\\shpinst\\shpleft0\\shptop0\\shpright2000\\shpbottom1000{\\sp{\\sn shapeType}{\\sv 3}}{\\sp{\\sn fillColor}{\\sv 255}}}}';
const TEXT_BOX = '{\\shp{\\*\\shpinst\\shpleft0\\shptop0\\shpright2000\\shpbottom1000{\\sp{\\sn shapeType}{\\sv 202}}{\\shptxt In the box\\par}}}';
const PICTURE =
  '{\\pict\\pngblip\\picw10\\pich10\\picwgoal150\\pichgoal150 89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d49444154789c6360000002000001e221bc330000000049454e44ae426082}';

test('convert: a document that loses nothing is converted without warnings', async () => {
  for (const bytes of [rtf('Hello {\\b rich} text'), rtf(`A picture ${PICTURE} in the text`), simpleDoc('First paragraph\rSecond paragraph\r').build()]) {
    const file = await open(bytes);
    const converted = await convertForEditing(file);
    assert.ok(converted);
    assert.deepEqual([converted.from, converted.to], [file.format, 'docx']);
    assert.deepEqual(converted.warnings, []);
    const again = await open(converted.bytes);
    assert.equal(again.format, 'docx');
    await again.document!();
  }
});

test('convert: shapes of an RTF document are counted in the warnings', async () => {
  const converted = await convertForEditing(await open(rtf(`Before ${SHAPE} between ${TEXT_BOX} after`)));
  assert.deepEqual(converted!.warnings, ['2 shape(s) were written as plain rectangles: their outline, fill and geometry were not carried over']);
});

test('convert: characters a docx cannot hold and pictures without data are counted in the warnings', async () => {
  const file = await open(rtf(`Text ${PICTURE} here`));
  const model = await file.document!();
  // What a reader may leave in the model: control characters in the text, a picture whose bytes it could not take.
  const paragraph = model.body.blocks[0] as Paragraph;
  const run = paragraph.children.find((inline) => inline.type === 'run' && inline.content[0].type === 'text');
  assert.ok(run && run.type === 'run' && run.content[0].type === 'text');
  run.content[0] = { type: 'text', text: 'Te\u0001xt\u0008 ' };
  assert.equal(model.media!.size, 1);
  model.media!.clear();
  const converted = await convertForEditing(file);
  assert.deepEqual(converted!.warnings, [
    '1 picture(s) whose data is not available were written as empty frames',
    '2 character(s) that a docx file cannot hold were left out',
  ]);
});

test('convert: an xls workbook reports the sheets whose content is not in the model; docx and xlsx are not converted', async () => {
  const sheet = cat(number(0, 0, 41));
  const plain = await convertForEditing(await open(xlsFile(workbookStream({ sheets: [{ name: 'Data', records: sheet }] }))));
  assert.deepEqual([plain!.from, plain!.to, plain!.warnings], ['xls', 'xlsx', []]);
  const charted = await convertForEditing(await open(xlsFile(workbookStream({ sheets: [{ name: 'Data', records: sheet }, { name: 'Chart1', type: 2 }] }))));
  assert.equal(charted!.warnings.length, 1);
  assert.match(charted!.warnings[0], /chart sheet/);

  assert.equal(await convertForEditing(await open(plain!.bytes)), undefined);
  const docx = await convertForEditing(await open(rtf('text')));
  assert.equal(await convertForEditing(await open(docx!.bytes)), undefined);
});
