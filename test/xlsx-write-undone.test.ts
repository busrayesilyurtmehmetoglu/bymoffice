// A sheet stays marked as modified once it was edited, also after the edit
// is undone. The writer does not take the mark for a change: a part that
// comes out as the bytes of the file asks for no new calculation, and the
// file saved after undo is the file that was opened.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { writeXlsx } from '../src/formats/xlsx/write';
import { ZipArchive, open } from '../src/index';
import { History, type Op } from '../src/ops/ops';
import { buildXlsx, worksheetXml } from './fixtures';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const decoder = new TextDecoder();
const encoder = new TextEncoder();
const same = (a: Uint8Array, b: Uint8Array): boolean => a.length === b.length && a.every((value, i) => value === b[i]);

async function book(calcChain = false, options: { tail?: string; target?: string; entry?: string } = {}): Promise<Uint8Array> {
  const bytes = await buildXlsx({
    sheets: [
      { name: 'S', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row></sheetData>') },
      { name: 'T', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><f>S!A1</f><v>1</v></c></row></sheetData>') },
    ],
    workbookTail: options.tail,
  });
  if (!calcChain) return bytes;
  const zip = ZipArchive.read(bytes);
  if (options.target !== undefined && options.entry === undefined) {
    // A relationship that calls another part of the workbook its chain.
    const rels = 'xl/_rels/workbook.xml.rels';
    zip.set(rels, encoder.encode(decoder.decode(await zip.read(rels)).replace('</Relationships>', `<Relationship Id="rIdChain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="${options.target}"/></Relationships>`)));
    return zip.write();
  }
  zip.set(options.entry ?? 'xl/calcChain.xml', encoder.encode(`${XML}<calcChain xmlns="${NS}"><c r="B1" i="1"/><c r="A1" i="2"/></calcChain>`));
  const rels = 'xl/_rels/workbook.xml.rels';
  zip.set(rels, encoder.encode(decoder.decode(await zip.read(rels)).replace('</Relationships>', '<Relationship Id="rIdChain" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>'.replace('calcChain.xml', options.target ?? 'calcChain.xml'))));
  const types = '[Content_Types].xml';
  zip.set(types, encoder.encode(decoder.decode(await zip.read(types)).replace('</Types>', '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>')));
  // Read again from the bytes it is saved as: what "the file" is from here on.
  return zip.write();
}

const EDITS: [name: string, ops: Op[]][] = [
  ['rows put in', [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }]],
  ['a column put in', [{ type: 'insert', sheet: 0, axis: 'col', at: 0, count: 1 }]],
  ['a row taken out', [{ type: 'delete', sheet: 0, axis: 'row', at: 0, count: 1 }]],
] as [string, Op[]][];

for (const [name, ops] of EDITS) {
  test(`${name}, saved, undone and saved again: the file is the one that was opened; redone, the one saved first`, async () => {
    const bytes = await book();
    const file = await open(bytes);
    const workbook = await file.workbook();
    const history = new History(workbook);
    history.apply(ops, undefined, undefined);
    await writeXlsx(file.package!, workbook);
    const edited = await file.package!.save();
    assert.ok(!same(edited, bytes), 'the edit is in the file');
    assert.ok(decoder.decode(await ZipArchive.read(edited).read('xl/workbook.xml')).includes('fullCalcOnLoad'), 'an edit asks for a new calculation');

    history.undo();
    await writeXlsx(file.package!, workbook);
    assert.ok(same(await file.package!.save(), bytes), 'undone: byte for byte the file that was opened');

    history.redo();
    await writeXlsx(file.package!, workbook);
    assert.ok(same(await file.package!.save(), edited), 'redone: byte for byte the file saved first');
  });
}

test('an edit that is undone before the first save leaves the file as it was', async () => {
  const bytes = await book();
  const file = await open(bytes);
  const workbook = await file.workbook();
  const history = new History(workbook);
  history.apply(EDITS[0][1], undefined, undefined);
  history.undo();
  assert.equal(workbook.sheets[0].modified, true, 'the mark stays; it is not what the writer goes by');
  await writeXlsx(file.package!, workbook);
  assert.ok(same(await file.package!.save(), bytes), 'byte for byte the file that was opened');
});

test('one sheet edited and undone beside one that stays edited: the file still asks for a new calculation', async () => {
  const bytes = await book();
  const file = await open(bytes);
  const workbook = await file.workbook();
  const history = new History(workbook);
  history.apply([{ type: 'insert', sheet: 1, axis: 'row', at: 0, count: 1 }] as Op[], undefined, undefined);
  history.apply(EDITS[0][1], undefined, undefined);
  history.undo();
  await writeXlsx(file.package!, workbook);
  const saved = ZipArchive.read(await file.package!.save());
  assert.ok(decoder.decode(await saved.read('xl/workbook.xml')).includes('fullCalcOnLoad'), 'the other sheet did change');
  assert.ok(same(await saved.read('xl/worksheets/sheet1.xml'), await ZipArchive.read(bytes).read('xl/worksheets/sheet1.xml')), 'the sheet that was undone is as the file had it');
});

test('the calculation chain goes with an edit and comes back, with the file as it was, when the edit is undone', async () => {
  const bytes = await book(true);
  const file = await open(bytes);
  const workbook = await file.workbook();
  const history = new History(workbook);
  history.apply(EDITS[0][1], undefined, undefined);
  await writeXlsx(file.package!, workbook);
  const edited = await file.package!.save();
  assert.ok(!ZipArchive.read(edited).has('xl/calcChain.xml'), 'a chain that no longer fits is taken out');
  history.undo();
  await writeXlsx(file.package!, workbook);
  assert.ok(same(await file.package!.save(), bytes), 'undone: the chain is back and the file is byte for byte the one that was opened');
});

test('undone before the first save, the calculation chain is never taken out', async () => {
  const bytes = await book(true);
  const file = await open(bytes);
  const workbook = await file.workbook();
  const history = new History(workbook);
  history.apply(EDITS[1][1], undefined, undefined);
  history.undo();
  await writeXlsx(file.package!, workbook);
  assert.ok(same(await file.package!.save(), bytes), 'byte for byte the file that was opened');
});

// ------------------------------------------------------------------ review

const NAMES = '<definedNames><definedName name="Far">S!$A$10</definedName><definedName name="_xlnm.Print_Area" localSheetId="0">S!$A$1:$C$50</definedName></definedNames>';

test('review: rows put in below the cells move what the defined names point at, and the file says so although the sheet part did not change', async () => {
  const bytes = await book(true, { tail: NAMES });
  const file = await open(bytes);
  const workbook = await file.workbook();
  const history = new History(workbook);
  history.apply([{ type: 'insert', sheet: 0, axis: 'row', at: 4, count: 2 }] as Op[], undefined, undefined);
  assert.deepEqual(workbook.definedNames.map((name) => name.value), ['S!$A$12', 'S!$A$1:$C$52']);
  await writeXlsx(file.package!, workbook);
  const saved = await file.package!.save();
  const archive = ZipArchive.read(saved);
  assert.ok(same(await archive.read('xl/worksheets/sheet1.xml'), await ZipArchive.read(bytes).read('xl/worksheets/sheet1.xml')), 'the part of the sheet is as the file had it');
  const again = await (await open(saved)).workbook();
  assert.deepEqual(again.definedNames.map((name) => name.value), ['S!$A$12', 'S!$A$1:$C$52'], 'the file read again names what the model names');
  assert.ok(decoder.decode(await archive.read('xl/workbook.xml')).includes('fullCalcOnLoad'), 'what reads the names is calculated again');
  assert.ok(!archive.has('xl/calcChain.xml'), 'and the chain is out');

  history.undo();
  await writeXlsx(file.package!, workbook);
  assert.ok(same(await file.package!.save(), bytes), 'undone: byte for byte the file that was opened');
});

test('review: a row taken out below the cells does the same', async () => {
  const bytes = await book(false, { tail: NAMES });
  const file = await open(bytes);
  const workbook = await file.workbook();
  new History(workbook).apply([{ type: 'delete', sheet: 0, axis: 'row', at: 4, count: 2 }] as Op[], undefined, undefined);
  await writeXlsx(file.package!, workbook);
  const again = await (await open(await file.package!.save())).workbook();
  assert.deepEqual(again.definedNames.map((name) => name.value), workbook.definedNames.map((name) => name.value));
  assert.equal(again.definedNames[0].value, 'S!$A$8');
});

test('review: a chain the relationship names in other letters than the archive comes back as the entry it was', async () => {
  const bytes = await book(true, { target: 'CalcChain.xml', entry: 'xl/calcChain.xml' });
  const file = await open(bytes);
  const workbook = await file.workbook();
  const history = new History(workbook);
  history.apply(EDITS[0][1], undefined, undefined);
  await writeXlsx(file.package!, workbook);
  await file.package!.save();
  history.undo();
  await writeXlsx(file.package!, workbook);
  assert.ok(same(await file.package!.save(), bytes), 'byte for byte the file that was opened');
});

for (const target of ['worksheets/sheet2.xml', 'workbook.xml']) {
  test(`review: a relationship that calls ${target} the calculation chain does not cost the workbook that part`, async () => {
    const bytes = await book(true, { target });
    const file = await open(bytes);
    const workbook = await file.workbook();
    new History(workbook).apply(EDITS[0][1], undefined, undefined);
    await writeXlsx(file.package!, workbook);
    const again = await (await open(await file.package!.save())).workbook();
    assert.deepEqual(again.sheets.map((sheet) => sheet.name), ['S', 'T'], 'the file opens with its sheets');
  });
}
