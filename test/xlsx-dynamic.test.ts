// Dynamic array formulas in xlsx files: read with their cell metadata, written back as Excel
// writes them, and a file that is not changed (or changed and undone) stays the file it was.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { needsDynamic } from '../src/calc/dynamic';
import { dynamicArrayCells, dynamicArrayMetadata, newMetadataXml } from '../src/formats/xlsx/metadata';
import { writeXlsx } from '../src/formats/xlsx/write';
import { createBlank } from '../src/formats/blank';
import { ZipArchive, getCell, open } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { History } from '../src/ops/ops';
import { buildXlsx, worksheetXml } from './fixtures';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const decoder = new TextDecoder();
const encoder = new TextEncoder();
const same = (a: Uint8Array, b: Uint8Array): boolean => a.length === b.length && a.every((value, i) => value === b[i]);
const text = async (bytes: Uint8Array, part: string): Promise<string> => decoder.decode(await ZipArchive.read(bytes).read(part));

/** The part Excel writes for a workbook with dynamic array formulas. */
const METADATA =
  `${XML}<metadata xmlns="${NS}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">` +
  '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>' +
  '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>' +
  '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';

/** A1 holds 3; B1 a dynamic array formula that spilled into B1:B3; C1 reads the spill. */
async function spilled(): Promise<Uint8Array> {
  return buildXlsx({
    sheets: [
      {
        name: 'S',
        xml: worksheetXml(
          '<sheetData>' +
            '<row r="1"><c r="A1"><v>3</v></c><c r="B1" cm="1"><f t="array" ref="B1:B3">_xlfn.SEQUENCE(A1)</f><v>1</v></c><c r="C1" cm="1"><f t="array" ref="C1">SUM(_xlfn.ANCHORARRAY(B1))</f><v>6</v></c></row>' +
            '<row r="2"><c r="B2"><v>2</v></c></row>' +
            '<row r="3"><c r="B3"><v>3</v></c></row>' +
            '</sheetData>',
        ),
      },
    ],
    metadata: METADATA,
  });
}

test('the metadata part: which values of cm mark dynamic array formulas', () => {
  assert.deepEqual(Array.from(dynamicArrayCells(encoder.encode(METADATA))), [1]);
  assert.deepEqual(Array.from(dynamicArrayCells(encoder.encode(newMetadataXml()))), [1]);
  // fDynamic="0", another type, a damaged part: none.
  assert.equal(dynamicArrayCells(encoder.encode(METADATA.replace('fDynamic="1"', 'fDynamic="0"'))).size, 0);
  assert.equal(dynamicArrayCells(encoder.encode(METADATA.replace(/XLDAPR/g, 'XLRICHVALUE'))).size, 0);
  assert.equal(dynamicArrayCells(encoder.encode('<metadata')).size, 0);
  assert.deepEqual(dynamicArrayMetadata(encoder.encode(METADATA)), { cm: 1 });
});

test('a metadata part of rich values gets a block for dynamic arrays after its own', () => {
  const rich =
    `${XML}<metadata xmlns="${NS}" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">` +
    '<metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/></metadataTypes>' +
    '<futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata>' +
    '<valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>';
  const made = dynamicArrayMetadata(encoder.encode(rich));
  assert.ok(made && made.xml);
  assert.equal(made.cm, 1);
  const xml = made.xml as string;
  // The rich value keeps its type (1) and its block; the new type is the second.
  assert.match(xml, /<metadataTypes count="2"><metadataType name="XLRICHVALUE"[^>]*\/><metadataType name="XLDAPR"/);
  assert.match(xml, /<\/futureMetadata><futureMetadata name="XLDAPR" count="1">.*<\/futureMetadata><cellMetadata count="1"><bk><rc t="2" v="0"\/><\/bk><\/cellMetadata><valueMetadata count="1"><bk><rc t="1" v="0"\/>/);
  // A part that already has cell metadata of another kind gets one block more.
  const other = METADATA.replace(/XLDAPR/g, 'XLOTHER');
  const added = dynamicArrayMetadata(encoder.encode(other));
  assert.ok(added && added.xml);
  assert.equal(added.cm, 2);
  assert.deepEqual(Array.from(dynamicArrayCells(encoder.encode(added.xml as string))), [2]);
  // A part in UTF-16 by its declaration is not added to.
  assert.equal(dynamicArrayMetadata(encoder.encode(other.replace('encoding="UTF-8"', 'encoding="UTF-16"'))), undefined);
});

test('a dynamic array formula of a file is read as one: an ordinary formula that knows where it spills', async () => {
  const file = await open(await spilled());
  const workbook = await file.workbook();
  const anchor = getCell(workbook.sheets[0], 0, 1)!;
  assert.deepEqual({ ...anchor.formula, attrs: undefined }, { text: '_xlfn.SEQUENCE(A1)', type: 'normal', ref: 'B1:B3', dynamic: true, attrs: undefined });
  assert.equal(anchor.cm, 1);
  // The engine trusts the values of the file, and recalculates what a change reaches.
  const engine = new CalcEngine(workbook);
  engine.recalculate();
  assert.equal(engine.unsupported.length, 0);
  assert.equal(engine.mismatched.length, 0);
});

test('opened and saved, or edited, undone and saved: the file is the one that was opened', async () => {
  const bytes = await spilled();
  const file = await open(bytes);
  const workbook = await file.workbook();
  const engine = new CalcEngine(workbook);
  engine.recalculate();
  await writeXlsx(file.package!, workbook);
  assert.ok(same(await file.package!.save(), bytes), 'not edited');

  const history = new History(workbook);
  history.apply(setCellInput(workbook, 0, 0, 0, '5'), undefined, undefined);
  engine.recalculate([{ sheet: 0, row: 0, col: 0 }]);
  const column = () => [0, 1, 2, 3, 4].map((row) => getCell(workbook.sheets[0], row, 1)?.value);
  assert.deepEqual(column(), [1, 2, 3, 4, 5]);
  assert.equal(getCell(workbook.sheets[0], 0, 2)?.value, 15);
  await writeXlsx(file.package!, workbook);
  const edited = await file.package!.save();
  const sheet = await text(edited, 'xl/worksheets/sheet1.xml');
  assert.match(sheet, /<c r="B1" cm="1"><f t="array" ref="B1:B5">_xlfn\.SEQUENCE\(A1\)<\/f><v>1<\/v><\/c>/);
  assert.match(sheet, /<c r="B5"><v>5<\/v><\/c>/);
  assert.equal(await text(edited, 'xl/metadata.xml'), METADATA, 'the metadata part of the file is reused as it is');

  history.undo();
  engine.recalculate([{ sheet: 0, row: 0, col: 0 }]);
  assert.deepEqual(column(), [1, 2, 3, undefined, undefined]);
  await writeXlsx(file.package!, workbook);
  assert.equal(await text(await file.package!.save(), 'xl/worksheets/sheet1.xml'), await text(bytes, 'xl/worksheets/sheet1.xml'), 'undone: the sheet of the file');
  assert.ok(same(await file.package!.save(), bytes), 'undone: byte for byte the file that was opened');
});

test('a formula typed into a new workbook spills, and is saved as Excel saves it', async () => {
  const file = await open(await createBlank('xlsx'));
  const workbook = await file.workbook();
  const engine = new CalcEngine(workbook);
  const typed = (row: number, col: number, input: string) => {
    for (const op of setCellInput(workbook, 0, row, col, input)) {
      if (op.type === 'setCells') for (const change of op.cells) {
        const sheet = workbook.sheets[0];
        let line = sheet.rows.get(change.row);
        line = line ? { ...line, cells: line.cells.slice() } : { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] };
        if (change.cell) line.cells[change.col] = change.cell;
        sheet.rows.set(change.row, line);
        sheet.changedRows = (sheet.changedRows || new Set()).add(change.row);
        sheet.modified = true;
        if (change.row >= sheet.rowCount) sheet.rowCount = change.row + 1;
        if (change.col >= sheet.colCount) sheet.colCount = change.col + 1;
      }
    }
    engine.recalculate([{ sheet: 0, row, col }]);
  };
  typed(0, 0, '=SEQUENCE(3,2)');
  typed(0, 3, '=XLOOKUP(3,A1:A3,B1:B3)');
  typed(0, 4, '=LET(x,A1:A3,SUM(x))');
  typed(0, 5, '=A1#*10');
  typed(4, 0, '=A1+1');
  await writeXlsx(file.package!, workbook);
  const saved = await file.package!.save();
  const sheet = await text(saved, 'xl/worksheets/sheet1.xml');
  assert.match(sheet, /<c r="A1" cm="1"><f t="array" ref="A1:B3">_xlfn\.SEQUENCE\(3,2\)<\/f><v>1<\/v><\/c>/);
  // A dynamic array formula of one value is one too: its own cell is where it spills.
  assert.match(sheet, /<c r="D1" cm="1"><f t="array" ref="D1">_xlfn\.XLOOKUP\(3,A1:A3,B1:B3\)<\/f><v>4<\/v><\/c>/);
  assert.match(sheet, /<c r="E1" cm="1"><f t="array" ref="E1">_xlfn\.LET\(_xlpm\.x,A1:A3,SUM\(_xlpm\.x\)\)<\/f><v>9<\/v><\/c>/);
  assert.match(sheet, /<c r="F1" cm="1"><f t="array" ref="F1:G3">_xlfn\.ANCHORARRAY\(A1\)\*10<\/f><v>10<\/v><\/c>/);
  assert.match(sheet, /<c r="A5"><f>A1\+1<\/f><v>2<\/v><\/c>/);
  assert.match(await text(saved, 'xl/metadata.xml'), /<metadataType name="XLDAPR"/);
  assert.match(await text(saved, '[Content_Types].xml'), /<Override PartName="\/xl\/metadata.xml" ContentType="application\/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata\+xml"\/>/);
  assert.match(await text(saved, 'xl/_rels/workbook.xml.rels'), /relationships\/sheetMetadata" Target="metadata.xml"/);

  // Read again: the same formulas, the same spills.
  const again = await (await open(saved)).workbook();
  const anchor = getCell(again.sheets[0], 0, 0)!;
  assert.equal(anchor.formula?.dynamic, true);
  assert.equal(anchor.formula?.ref, 'A1:B3');
  assert.equal(getCell(again.sheets[0], 2, 1)?.value, 6);
  assert.equal(getCell(again.sheets[0], 0, 4)?.formula?.dynamic, true, 'a dynamic formula of one value stays one');
  assert.equal(getCell(again.sheets[0], 4, 0)?.formula?.dynamic, undefined, 'an ordinary formula that means the same');
});

test('whether a formula is a dynamic one is a matter of its text, not of what it is worth at the moment', async () => {
  const file = await open(await createBlank('xlsx'));
  const workbook = await file.workbook();
  const engine = new CalcEngine(workbook);
  const history = new History(workbook);
  const typed = (row: number, col: number, input: string) => {
    history.apply(setCellInput(workbook, 0, row, col, input), undefined, undefined);
    engine.recalculate([{ sheet: 0, row, col }]);
  };
  typed(0, 2, '1');
  typed(1, 2, '7');
  typed(2, 2, '3');
  // One value today, several once more of the numbers are large.
  typed(0, 0, '=FILTER(C1:C3,C1:C3>5)');
  assert.equal(getCell(workbook.sheets[0], 0, 0)?.value, 7);
  await writeXlsx(file.package!, workbook);
  const saved = await file.package!.save();
  assert.match(await text(saved, 'xl/worksheets/sheet1.xml'), /<c r="A1" cm="1"><f t="array" ref="A1">_xlfn\._xlws\.FILTER\(C1:C3,C1:C3&gt;5\)<\/f><v>7<\/v><\/c>/);
  const again = await (await open(saved)).workbook();
  assert.equal(getCell(again.sheets[0], 0, 0)?.formula?.dynamic, true);
  const later = new CalcEngine(again);
  const row = again.sheets[0].rows.get(0)!;
  again.sheets[0].rows.set(0, { ...row, cells: Object.assign(row.cells.slice(), { 2: { kind: 'number', value: 9, style: 0 } }) });
  later.recalculate([{ sheet: 0, row: 0, col: 2 }]);
  assert.deepEqual([0, 1].map((r) => getCell(again.sheets[0], r, 0)?.value), [9, 7], 'reopened, it spills as it did before the save');
});

test('needsDynamic: what an ordinary formula would read otherwise, by its text', () => {
  const dynamic = ['A1:A3', 'A1:A3*2', 'SUM(A1:A3*2)', 'ABS(A1:A2)', 'IF(A1>0,B1:B3,0)', 'FILTER(A1:A3,A1:A3>1)', 'SUM(SEQUENCE(3))', '{1,2}', 'XLOOKUP(1,A:A,B:C)', 'Rate*2', 'LET(x,1,x)', 'A1#', 'INDEX(A1:B3,2)', 'ROW(A1:A3)', 'OFFSET(A1,0,0,3)', 'NOSUCH(A1)'];
  const ordinary = ['A1+1', 'SUM(A1:A3)', 'SUM(A:A)', 'SUMIF(A1:A9,">1",B1:B9)', 'VLOOKUP(1,A1:C9,2,FALSE)', 'SUMPRODUCT(A1:A3*2)', 'INDEX(A1:B3,2,1)', '@A1:A3*2', 'SUM(OFFSET(A1,0,0,3))', 'ROW()', 'ROWS(A1:A3)', 'COUNTIF(A:A,1)', 'A1&"x"', 'AND(A1:A3)', 'SUM(@SEQUENCE(3))'];
  for (const text of dynamic) assert.equal(needsDynamic(text), true, text);
  for (const text of ordinary) assert.equal(needsDynamic(text), false, text);
});

test('a metadata part of very many blocks, or one with comments, is read without trouble', () => {
  const blocks = METADATA.replace('<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>', `<cellMetadata count="200000">${'<bk><rc t="1" v="0"/></bk>'.repeat(200000)}</cellMetadata>`);
  assert.deepEqual(dynamicArrayMetadata(encoder.encode(blocks)), { cm: 1 });
  // What a comment or a CDATA section holds is not taken for an element: the part is added to around it, and keeps it (K2).
  const commented = METADATA.replace(/XLDAPR/g, 'XLOTHER').replace('<metadataTypes', '<!-- <cellMetadata count="9"> --><metadataTypes');
  const made = dynamicArrayMetadata(encoder.encode(commented));
  assert.ok(made && made.xml);
  assert.equal(made.cm, 2);
  assert.ok(made.xml.indexOf('<!-- <cellMetadata count="9"> -->') >= 0);
  assert.equal(dynamicArrayCells(encoder.encode(made.xml)).has(2), true);
  assert.equal(dynamicArrayCells(encoder.encode(made.xml)).has(1), false);
  const cdata = commented.replace('<!-- <cellMetadata count="9"> -->', '').replace('</metadata>', '<extLst><ext uri="{x}"><x:y xmlns:x="urn:x"><![CDATA[<cellMetadata count="7">]]></x:y></ext></extLst></metadata>');
  const withCdata = dynamicArrayMetadata(encoder.encode(cdata));
  assert.ok(withCdata && withCdata.xml && withCdata.xml.indexOf('<![CDATA[<cellMetadata count="7">]]>') >= 0);
  assert.equal(dynamicArrayCells(encoder.encode(withCdata.xml)).has(withCdata.cm), true);
  // Very many comments are looked through once (before: each comment searched the rest of the part again, 26 seconds).
  const many = METADATA.replace(/XLDAPR/g, 'XLOTHER').replace('<metadataTypes', '<!---->'.repeat(320000) + '<metadataTypes');
  const started = Date.now();
  const lots = dynamicArrayMetadata(encoder.encode(many));
  assert.ok(lots && lots.cm === 2);
  assert.ok(Date.now() - started < 3000, `${Date.now() - started} ms`);
  // A comment that never ends: the part is not changed, and the formula is written without a mark.
  assert.equal(dynamicArrayMetadata(encoder.encode(commented.replace(' --><metadataTypes', '<metadataTypes'))), undefined);
});

test('a part called xl/metadata.xml that nothing leads to stays; the new one gets another name', async () => {
  const zip = ZipArchive.read(await buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml('<sheetData/>') }] }));
  zip.set('xl/metadata.xml', encoder.encode('<orphan>keep me</orphan>'));
  const file = await open(await zip.write());
  const workbook = await file.workbook();
  const engine = new CalcEngine(workbook);
  new History(workbook).apply(setCellInput(workbook, 0, 0, 0, '=SEQUENCE(3)'), undefined, undefined);
  engine.recalculate([{ sheet: 0, row: 0, col: 0 }]);
  await writeXlsx(file.package!, workbook);
  const saved = await file.package!.save();
  assert.equal(await text(saved, 'xl/metadata.xml'), '<orphan>keep me</orphan>');
  assert.match(await text(saved, 'xl/metadata2.xml'), /<metadataType name="XLDAPR"/);
  assert.match(await text(saved, 'xl/_rels/workbook.xml.rels'), /relationships\/sheetMetadata" Target="metadata2.xml"/);
  assert.match(await text(saved, '[Content_Types].xml'), /PartName="\/xl\/metadata2.xml"/);
});
