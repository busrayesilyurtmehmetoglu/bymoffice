import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CfbFile } from '../src/cfb';
import { RecordReader, readerAt, recordStream } from '../src/formats/xls/records';
import { decodeBytes } from '../src/formats/xls/codepage';
import { decodeRk, readXls, readXlsStream, readXlsWithReport } from '../src/formats/xls/workbook';
import { getCell } from '../src/model/sheet/types';
import {
  blank,
  bof,
  boolErr,
  cat,
  eof,
  f64,
  font,
  formula,
  labelSst,
  narrow,
  number,
  ptg,
  rec,
  rk,
  rkFloat,
  rkInt,
  rkInt100,
  sst,
  str16,
  u16,
  u32,
  wide,
  workbookStream,
  xf,
  xlsFile,
} from './xls-fixtures';

const one = (records: Uint8Array, tail?: Uint8Array) => readXlsStream(workbookStream({ sheets: [{ name: 'Sheet1', records }], tail })).workbook;

test('record reader: values that cross into CONTINUE records, empty CONTINUE records included', () => {
  const bytes = cat(rec(0x1234, [1, 2, 3]), rec(0x003c), rec(0x003c, [4]), rec(0x003c, [5, 6, 7, 8, 9, 10, 11, 12, 13]), rec(0x0099, [99]));
  const r = readerAt(recordStream(bytes), 0);
  assert.equal(r.remaining, 13);
  assert.equal(r.u16(), 0x0201);
  assert.equal(r.u32(), 0x06050403);
  assert.equal(r.left, 7);
  r.skip(3);
  assert.deepEqual([...r.take(2)], [10, 11]);
  assert.equal(r.done, false);
  assert.equal(r.u16(), 0x0d0c);
  assert.equal(r.done, true);
  assert.throws(() => r.u8(), /Invalid xls/);
});

test('record reader: a double split between records', () => {
  const value = f64(-1234.5678);
  const bytes = cat(rec(0x0001, value.slice(0, 3)), rec(0x003c, value.slice(3)));
  assert.equal(readerAt(recordStream(bytes), 0).f64(), -1234.5678);
});

test('record reader: a record that runs past the stream is refused', () => {
  const bytes = cat(u16(0x0001), u16(50), [1, 2, 3]);
  assert.throws(() => readerAt(recordStream(bytes), 0), /runs past the end/);
  assert.ok(RecordReader);
});

test('RK: integer, integer in hundredths, double, double in hundredths', () => {
  assert.equal(decodeRk(rkInt(42)), 42);
  assert.equal(decodeRk(rkInt(-7) | 0), -7);
  assert.equal(decodeRk(rkInt100(12345)), 123.45);
  assert.equal(decodeRk(rkInt100(-250) | 0), -2.5);
  assert.equal(decodeRk(rkFloat(1.5)), 1.5);
  assert.equal(decodeRk(rkFloat(-3.75)), -3.75);
  assert.equal(decodeRk(rkFloat(25, true)), 0.25);
  assert.equal(decodeRk(rkFloat(2 ** 40)), 2 ** 40);
  // The largest and the smallest 30-bit integers.
  assert.equal(decodeRk(rkInt(536870911)), 536870911);
  assert.equal(decodeRk(rkInt(-536870912) | 0), -536870912);
});

test('cell records: NUMBER, RK, MULRK, BLANK, MULBLANK, BOOLERR', () => {
  const workbook = one(
    cat(
      number(0, 0, 3.14159),
      rk(0, 1, rkInt100(1999)),
      rec(0x00bd, u16(1), u16(2), u16(15), u32(rkInt(1)), u16(15), u32(rkInt(2)), u16(15), u32(rkFloat(0.5)), u16(4)),
      blank(2, 0, 15),
      rec(0x00be, u16(3), u16(1), u16(15), u16(15), u16(15), u16(3)),
      boolErr(4, 0, 1, false),
      boolErr(4, 1, 0, false),
      boolErr(4, 2, 0x07, true),
      boolErr(4, 3, 0x2a, true),
    ),
  );
  const sheet = workbook.sheets[0];
  assert.deepEqual(getCell(sheet, 0, 0), { kind: 'number', value: 3.14159, style: 0 });
  assert.equal(getCell(sheet, 0, 1)?.value, 19.99);
  assert.deepEqual([2, 3, 4].map((col) => getCell(sheet, 1, col)?.value), [1, 2, 0.5]);
  assert.equal(getCell(sheet, 1, 5), undefined);
  assert.deepEqual(getCell(sheet, 2, 0), { kind: 'blank', value: null, style: 0 });
  assert.deepEqual([0, 1, 2, 3, 4].map((col) => getCell(sheet, 3, col)?.kind), [undefined, 'blank', 'blank', 'blank', undefined]);
  assert.deepEqual(getCell(sheet, 4, 0), { kind: 'boolean', value: true, style: 0 });
  assert.equal(getCell(sheet, 4, 1)?.value, false);
  assert.deepEqual(getCell(sheet, 4, 2), { kind: 'error', value: '#DIV/0!', style: 0 });
  assert.equal(getCell(sheet, 4, 3)?.value, '#N/A');
  assert.equal(sheet.rowCount, 5);
  assert.equal(sheet.colCount, 5);
});

test('a cell written twice: the later record wins', () => {
  const sheet = one(cat(number(0, 0, 1), number(0, 0, 2))).sheets[0];
  assert.equal(getCell(sheet, 0, 0)?.value, 2);
});

test('a number that is not finite reads as #NUM!', () => {
  const sheet = one(number(0, 0, Infinity)).sheets[0];
  assert.deepEqual(getCell(sheet, 0, 0), { kind: 'error', value: '#NUM!', style: 0 });
});

test('LABELSST: shared strings, compressed and UTF-16', () => {
  const workbook = one(cat(labelSst(0, 0, 0), labelSst(0, 1, 1), labelSst(1, 0, 1)), sst(str16('Fiyat'), str16('Çığ şöleni İĞÜ')));
  const sheet = workbook.sheets[0];
  assert.deepEqual(getCell(sheet, 0, 0), { kind: 'string', value: 'Fiyat', style: 0, si: 0 });
  assert.equal(getCell(sheet, 0, 1)?.value, 'Çığ şöleni İĞÜ');
  assert.equal(getCell(sheet, 1, 0)?.si, 1);
  assert.deepEqual(workbook.sharedStrings, [{ text: 'Fiyat' }, { text: 'Çığ şöleni İĞÜ' }]);
});

test('SST: a string split by a CONTINUE record, with the character width changing at the split', () => {
  // "abc" + "çığ": the first part one byte per character, the rest two.
  const first = cat(u32(2), u32(2), u16(6), [0], narrow('abc'));
  const second = cat([1], wide('çığ'), str16('son'));
  const workbook = one(cat(labelSst(0, 0, 0), labelSst(0, 1, 1)), cat(rec(0x00fc, first), rec(0x003c, second)));
  assert.equal(getCell(workbook.sheets[0], 0, 0)?.value, 'abcçığ');
  assert.equal(getCell(workbook.sheets[0], 0, 1)?.value, 'son');
});

test('SST: wide to compressed at the split, an empty CONTINUE in between, and a string that starts a CONTINUE', () => {
  const first = cat(u32(3), u32(3), u16(5), [1], wide('ğü'));
  const second = cat([0], narrow('xyz'));
  // A string that starts a CONTINUE record has no extra flags byte before its header.
  const third = cat(str16('tam'), u16(2), [0]);
  const fourth = cat([0], narrow('ok'));
  const tail = cat(rec(0x00fc, first), rec(0x003c), rec(0x003c, second), rec(0x003c, third), rec(0x003c, fourth));
  const workbook = one(cat(labelSst(0, 0, 0), labelSst(0, 1, 1), labelSst(0, 2, 2)), tail);
  assert.deepEqual([0, 1, 2].map((col) => getCell(workbook.sheets[0], 0, col)?.value), ['ğüxyz', 'tam', 'ok']);
});

test('SST: rich text runs become runs of the model, phonetic data is skipped', () => {
  const head = cat(font(), font({ bold: true }), font({ italic: true }), font({ size: 14 }), font({ name: 'Courier', color: 10 }), font({ strike: true }));
  const xfs: Uint8Array[] = [];
  for (let i = 0; i < 15; i++) xfs.push(xf({ style: true }));
  xfs.push(xf());
  // "Kalın şeker": bold from 0 (font 1), font index 5 (the fifth record: the index 4 is skipped) from 5.
  const rich = cat(u16(11), [0x01 | 0x08 | 0x04], u16(2), u32(6), wide('Kalın şeker'), u16(0), u16(1), u16(5), u16(5), [9, 9, 9, 9, 9, 9]);
  const plain = str16('düz');
  // A run that starts after the first character leaves the start in the font of the cell.
  const partly = cat(u16(4), [0x08], u16(1), narrow('abcd'), u16(2), u16(2));
  const stream = workbookStream({
    head: cat(head, ...xfs),
    tail: sst(rich, plain, partly),
    sheets: [{ name: 'S', records: cat(labelSst(0, 0, 0), labelSst(0, 1, 1), labelSst(0, 2, 2)) }],
  });
  const workbook = readXlsStream(stream).workbook;
  const cell = getCell(workbook.sheets[0], 0, 0);
  assert.equal(cell?.value, 'Kalın şeker');
  assert.equal(cell?.runs?.length, 2);
  assert.deepEqual(cell?.runs?.map((run) => run.text), ['Kalın', ' şeker']);
  assert.equal(cell?.runs?.[0].font?.bold, true);
  assert.equal(cell?.runs?.[1].font?.name, 'Courier');
  assert.deepEqual(cell?.runs?.[1].font?.color, { indexed: 10 });
  assert.equal(getCell(workbook.sheets[0], 0, 1)?.value, 'düz');
  assert.equal(getCell(workbook.sheets[0], 0, 1)?.runs, undefined);
  const third = getCell(workbook.sheets[0], 0, 2);
  assert.deepEqual(third?.runs?.map((run) => [run.text, run.font?.italic]), [['ab', undefined], ['cd', true]]);
  assert.equal(workbook.sharedStrings[0].runs, cell?.runs);
});

test('SST: a count larger than the strings present does not allocate or fail', () => {
  const tail = rec(0x00fc, u32(0xffffffff), u32(0xffffffff), str16('tek'));
  const workbook = one(labelSst(0, 0, 0), tail);
  assert.equal(workbook.sharedStrings.length, 1);
  assert.equal(getCell(workbook.sheets[0], 0, 0)?.value, 'tek');
});

test('LABELSST with an index outside the table is an error that names the cell', () => {
  assert.throws(() => one(labelSst(2, 1, 5), sst(str16('a'))), /shared string index 5 in cell B3 .* out of range \(1 strings\)/);
});

test('LABEL and RSTRING (BIFF8)', () => {
  const label = rec(0x0204, u16(0), u16(0), u16(15), str16('Etiket ğ'));
  const rstring = rec(0x00d6, u16(1), u16(0), u16(15), str16('abcdef'), u16(1), u16(3), u16(1));
  const sheet = one(cat(label, rstring)).sheets[0];
  assert.deepEqual(getCell(sheet, 0, 0), { kind: 'string', value: 'Etiket ğ', style: 0 });
  assert.deepEqual(getCell(sheet, 1, 0)?.runs?.map((run) => run.text), ['abc', 'def']);
});

test('FORMULA results: number, text (STRING record), boolean, error, empty text', () => {
  const two = [...ptg.int(1), ...ptg.int(1), ...ptg.add];
  const sheet = one(
    cat(
      formula(0, 0, two, 2),
      formula(0, 1, two, { text: 'Sonuç' }),
      formula(0, 2, two, true),
      formula(0, 3, two, { error: 0x0f }),
      formula(0, 4, two, 'emptyText'),
      // A STRING record with nothing waiting for it is ignored.
      rec(0x0207, str16('yetim')),
      formula(0, 5, two, { text: '' }),
    ),
  ).sheets[0];
  assert.deepEqual(getCell(sheet, 0, 0), { kind: 'number', value: 2, style: 0, formula: { text: '1+1', type: 'normal' } });
  assert.deepEqual(getCell(sheet, 0, 1), { kind: 'string', value: 'Sonuç', style: 0, formula: { text: '1+1', type: 'normal' } });
  assert.equal(getCell(sheet, 0, 2)?.value, true);
  assert.deepEqual([getCell(sheet, 0, 3)?.kind, getCell(sheet, 0, 3)?.value], ['error', '#VALUE!']);
  assert.deepEqual([getCell(sheet, 0, 4)?.kind, getCell(sheet, 0, 4)?.value], ['string', '']);
  assert.equal(getCell(sheet, 0, 5)?.value, '');
});

test('a formula text result continued in a CONTINUE record', () => {
  const two = [...ptg.int(1), ...ptg.int(1), ...ptg.add];
  const record = cat(formula(0, 0, two, { text: '' }).subarray(0, 4 + 22 + two.length), rec(0x0207, u16(6), [0], narrow('abc')), rec(0x003c, [1], wide('çığ')));
  assert.equal(getCell(one(record).sheets[0], 0, 0)?.value, 'abcçığ');
});

test('code pages: Turkish (1254) and the others carried in the reader agree with the platform decoder', () => {
  const bytes = new Uint8Array(256);
  for (let i = 0; i < 256; i++) bytes[i] = i;
  assert.equal(decodeBytes(new Uint8Array([0xd0, 0xf0, 0xdd, 0xfd, 0xde, 0xfe, 0xc7, 0xe7, 0xd6, 0xf6, 0xdc, 0xfc]), 0, 12, 1254), 'ĞğİıŞşÇçÖöÜü');
  for (const [codepage, label] of [[1250, 'windows-1250'], [1251, 'windows-1251'], [1253, 'windows-1253'], [1254, 'windows-1254']] as const) {
    let expected: string;
    try {
      expected = new TextDecoder(label).decode(bytes);
    } catch {
      continue;
    }
    const actual = decodeBytes(bytes, 0, 256, codepage);
    for (let i = 0; i < 256; i++) {
      // 0xAA has no character in the Greek code page; decoders disagree on what to make of it.
      if (codepage === 1253 && i === 0xaa) continue;
      assert.equal(actual.charCodeAt(i), expected.charCodeAt(i), `${label} byte 0x${i.toString(16)}`);
    }
  }
  // Windows-1252 is the Turkish page with eight letters different (some platform decoders read it as Latin-1, so it is not compared with them).
  const western = decodeBytes(bytes, 0, 256, 1252);
  const turkish = decodeBytes(bytes, 0, 256, 1254);
  const different: number[] = [];
  for (let i = 0; i < 256; i++) if (western[i] !== turkish[i]) different.push(i);
  assert.deepEqual(different, [0x8e, 0x9e, 0xd0, 0xdd, 0xde, 0xf0, 0xfd, 0xfe]);
  assert.equal(decodeBytes(new Uint8Array([0x80, 0x8e, 0x9e, 0x9f, 0xd0, 0xfe, 0xe9]), 0, 7, 1252), '€ŽžŸÐþé');
  // DOS code pages: spot checks.
  assert.equal(decodeBytes(new Uint8Array([0x80, 0x81, 0x9b, 0xe1, 0xf8, 0xff]), 0, 6, 437), 'Çü¢ß° ');
  assert.equal(decodeBytes(new Uint8Array([0x80, 0x9b, 0x9d, 0xd5, 0xf5]), 0, 5, 850), 'ÇøØı§');
  // UTF-8, with a malformed byte and a character outside the basic plane.
  assert.equal(decodeBytes(new Uint8Array([0x41, 0xc4, 0x9f, 0xe2, 0x82, 0xac, 0xf0, 0x9f, 0x98, 0x80, 0xff, 0xc0, 0x80]), 0, 13, 65001), 'Ağ€😀���');
  // An unknown code page reads as Windows-1252.
  assert.equal(decodeBytes(new Uint8Array([0xe9, 0x80]), 0, 2, 12345), 'é€');
});

test('BIFF5: byte strings in the code page of the workbook (LABEL, sheet name, font name, FORMULA text)', () => {
  const tr = [0xc7, 0x61, 0x79, 0x20, 0xfe, 0x65, 0x6b, 0x65, 0x72, 0x69]; // "Çay şekeri" in Windows-1254
  const label = rec(0x0204, u16(0), u16(0), u16(0), u16(tr.length), tr);
  const rstring = rec(0x00d6, u16(1), u16(0), u16(0), u16(4), narrow('abcd'), [1], [2, 1]);
  const text = cat(rec(0x0006, u16(2), u16(0), u16(0), [0, 0, 0, 0, 0, 0, 0xff, 0xff], u16(0), u32(0), u16(3), ptg.int(7)), rec(0x0207, u16(2), [0xd0, 0xf0]));
  const stream = cat(
    bof(0x0005, 0x0500),
    rec(0x0042, u16(1254)),
    rec(0x0031, u16(200), u16(0), u16(0x7fff), u16(400), u16(0), [0, 0, 162, 0], [3], [0xdd, 0x6e, 0x63]),
    rec(0x0031, u16(240), u16(0), u16(0x7fff), u16(700), u16(0), [0, 0, 162, 0], [1], narrow('B')),
    rec(0x00e0, u16(0), u16(0), u16(0x0001), [0x20, 0], u32(0), u32(0)),
    // The sheet offset is filled in below.
    rec(0x0085, u32(0), [0, 0], [3], [0xde, 0x75, 0x62]),
    eof(),
  );
  const offset = stream.length;
  const sheet = cat(bof(0x0010, 0x0500), label, rstring, text, eof());
  const bytes = cat(stream, sheet);
  // Patch the BOUNDSHEET offset: the record is the last but one of the globals.
  const boundAt = offset - 4 - (4 + 4 + 2 + 1 + 3) + 4;
  new DataView(bytes.buffer).setUint32(boundAt, offset, true);
  const { workbook, report } = readXlsStream(bytes, 'Book');
  assert.equal(report.biff, 5);
  assert.equal(report.codepage, 1254);
  assert.equal(workbook.sheets[0].name, 'Şub');
  assert.equal(workbook.styles.fonts[0].name, 'İnc');
  assert.equal(getCell(workbook.sheets[0], 0, 0)?.value, 'Çay şekeri');
  assert.deepEqual(getCell(workbook.sheets[0], 1, 0)?.runs?.map((run) => [run.text, run.font?.bold]), [['ab', undefined], ['cd', true]]);
  assert.deepEqual(getCell(workbook.sheets[0], 2, 0), { kind: 'string', value: 'Ğğ', style: 0, formula: { text: '7', type: 'normal' } });
});

test('readXls: through the compound file, and the workbook-level properties', () => {
  const stream = workbookStream({
    head: cat(rec(0x0042, u16(1200)), rec(0x0022, u16(1)), rec(0x003d, u16(0), u16(0), u16(100), u16(100), u16(0x38), u16(1), u16(0), u16(1), u16(600)), font(), xf({ style: true }), xf()),
    sheets: [
      { name: 'Veri', records: number(0, 0, 1) },
      { name: 'Gizli sayfa', visibility: 1, records: number(0, 0, 2) },
      { name: 'Çok gizli', visibility: 2 },
      { name: 'Grafik', type: 2, records: number(0, 0, 99) },
      { name: 'Makro', type: 1, records: number(0, 0, 98) },
    ],
  });
  const cfb = CfbFile.read(xlsFile(stream));
  const workbook = readXls(cfb);
  assert.deepEqual(workbook.sheets.map((sheet) => [sheet.name, sheet.state, sheet.kind, sheet.sheetId]), [
    ['Veri', 'visible', 'worksheet', 1],
    ['Gizli sayfa', 'hidden', 'worksheet', 2],
    ['Çok gizli', 'veryHidden', 'worksheet', 3],
    ['Grafik', 'visible', 'chartsheet', 4],
    ['Makro', 'visible', 'macrosheet', 5],
  ]);
  assert.equal(workbook.date1904, true);
  assert.equal(workbook.activeSheet, 1);
  assert.equal(getCell(workbook.sheets[1], 0, 0)?.value, 2);
  // Charts and macro sheets are not read: no cells, nothing run.
  assert.equal(workbook.sheets[3].rows.size, 0);
  assert.equal(workbook.sheets[4].rows.size, 0);
  assert.equal(readXlsWithReport(cfb).report.biff, 8);
});

test('a sheet whose BOF says chart or macro sheet stays empty whatever the sheet list says', () => {
  const stream = workbookStream({ sheets: [{ name: 'A', bofType: 0x0020, records: number(0, 0, 1) }, { name: 'B', bofType: 0x0040, records: number(0, 0, 1) }] });
  const workbook = readXlsStream(stream).workbook;
  assert.deepEqual(workbook.sheets.map((sheet) => [sheet.kind, sheet.rows.size]), [['chartsheet', 0], ['macrosheet', 0]]);
});

test('an embedded chart substream inside a sheet is skipped, the cells after it are read', () => {
  const chart = cat(bof(0x0020), number(5, 5, 555), bof(0x0020), eof(), eof());
  const sheet = one(cat(number(0, 0, 1), chart, number(1, 0, 2))).sheets[0];
  assert.equal(getCell(sheet, 0, 0)?.value, 1);
  assert.equal(getCell(sheet, 1, 0)?.value, 2);
  assert.equal(getCell(sheet, 5, 5), undefined);
});

test('records the reader does not know are counted, comments too', () => {
  const { report } = readXlsStream(workbookStream({ sheets: [{ name: 'S', records: cat(rec(0x00ec, [1, 2, 3]), rec(0x00ec), rec(0x001c, u16(0), u16(0), u16(0), u16(1), u16(0)), rec(0x1234)) }] }));
  // Drawing records are read now (see xls-drawings.test.ts); these two hold nothing that can be read.
  assert.equal(report.skippedRecords.get(0x00ec), undefined);
  assert.equal(report.skippedRecords.get(0x1234), 1);
  assert.equal(report.comments, 1);
});
