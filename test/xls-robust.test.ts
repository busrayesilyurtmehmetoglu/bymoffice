import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CfbFile } from '../src/cfb';
import { readXls, readXlsStream } from '../src/formats/xls/workbook';
import { getCell } from '../src/model/sheet/types';
import {
  arrayRec,
  blank,
  bof,
  cat,
  definedName,
  eof,
  externSheet,
  formula,
  labelSst,
  number,
  ptg,
  rec,
  shrFmla,
  sst,
  str16,
  str8,
  supBookSelf,
  tokens,
  u16,
  u32,
  wide,
  workbookStream,
  xlsFile,
} from './xls-fixtures';

const one = (records: Uint8Array, tail?: Uint8Array) => workbookStream({ sheets: [{ name: 'Sheet1', records }], tail });
const read = (records: Uint8Array, tail?: Uint8Array) => readXlsStream(one(records, tail)).workbook;

test('an encrypted workbook is refused with UnsupportedFileError', () => {
  const stream = workbookStream({ head: rec(0x002f, u16(1), u16(1), u16(1), new Array(48).fill(0)), sheets: [{ name: 'S' }] });
  assert.throws(() => readXlsStream(stream), (error: Error) => error.name === 'UnsupportedFileError' && /Password-protected/.test(error.message));
  assert.throws(() => readXls(CfbFile.read(xlsFile(stream))), /Password-protected/);
});

test('streams that are not a BIFF5 or BIFF8 workbook', () => {
  assert.throws(() => readXlsStream(new Uint8Array(0)), /Invalid xls: the workbook stream is empty/);
  assert.throws(() => readXlsStream(new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8])), /Invalid xls/);
  assert.throws(() => readXlsStream(cat(rec(0x0203, u16(0)), eof())), /does not start with a BOF record/);
  // BIFF2 to BIFF4 start with other BOF records.
  for (const id of [0x0009, 0x0209, 0x0409]) {
    assert.throws(() => readXlsStream(cat(rec(id, u16(0), u16(0x10)), eof())), (error: Error) => error.name === 'UnsupportedFileError');
  }
  // A stream that starts with a sheet instead of the workbook globals.
  assert.throws(() => readXlsStream(cat(bof(0x0010), eof())), (error: Error) => error.name === 'UnsupportedFileError');
  assert.throws(() => readXlsStream(cat(bof(0x0005), eof())), /has no sheets/);
});

test('a compound file without a workbook stream', () => {
  const cfb = { has: () => false, stream: () => new Uint8Array(0) } as unknown as CfbFile;
  assert.throws(() => readXls(cfb), (error: Error) => error.name === 'UnsupportedFileError');
});

test('truncated streams: globals or a sheet without its end, a record longer than the stream', () => {
  const stream = one(cat(number(0, 0, 1), number(1, 0, 2)));
  assert.throws(() => readXlsStream(cat(bof(0x0005), rec(0x0042, u16(1252)))), /globals are not terminated/);
  assert.throws(() => readXlsStream(stream.subarray(0, stream.length - 4)), /is not terminated/);
  assert.throws(() => readXlsStream(stream.subarray(0, stream.length - 10)), /runs past the end|not terminated/);
  assert.throws(() => readXlsStream(cat(bof(0x0005), u16(0x0031), u16(60000), [1, 2, 3])), /runs past the end/);
  // Every prefix of a workbook either reads or fails with an Error; none hangs.
  const rich = richWorkbook();
  for (let size = 0; size < rich.length; size += 3) {
    try {
      readXlsStream(rich.subarray(0, size));
    } catch (error) {
      assert.ok(error instanceof Error, `prefix of ${size} bytes`);
    }
  }
});

test('sheet offsets: outside the stream, not at a BOF record, shared by two sheets', () => {
  const patch = (stream: Uint8Array, offset: number) => {
    // The first BOUNDSHEET record follows the globals BOF and the default style records.
    const copy = stream.slice();
    const view = new DataView(copy.buffer);
    for (let p = 0; p < copy.length; p += 4 + view.getUint16(p + 2, true)) {
      if (view.getUint16(p, true) === 0x0085) {
        view.setUint32(p + 4, offset, true);
        return copy;
      }
    }
    throw new Error('no BOUNDSHEET');
  };
  const stream = one(number(0, 0, 1));
  assert.throws(() => readXlsStream(patch(stream, 0xfffffff0)), /does not start with a BOF record/);
  assert.throws(() => readXlsStream(patch(stream, stream.length - 2)), /does not start with a BOF record/);
  assert.throws(() => readXlsStream(patch(stream, 30)), /does not start with a BOF record/);
  // Pointing a sheet at the globals reads the globals as a sheet: nothing it holds is a cell.
  assert.doesNotThrow(() => readXlsStream(patch(stream, 0)));

  const two = workbookStream({ sheets: [{ name: 'A', records: number(0, 0, 1) }, { name: 'B', records: number(0, 0, 2) }] });
  const view = new DataView(two.buffer);
  const offsets: number[] = [];
  for (let p = 0; p < two.length; p += 4 + view.getUint16(p + 2, true)) if (view.getUint16(p, true) === 0x0085) offsets.push(p + 4);
  view.setUint32(offsets[1], view.getUint32(offsets[0], true), true);
  assert.throws(() => readXlsStream(two), /sheet "B" shares its records with another sheet/);
});

test('cells outside the sheet limits and records too short for their fields', () => {
  assert.throws(() => read(number(0, 256, 1)), /column 257, past the last column/);
  assert.throws(() => read(rec(0x00be, u16(0), u16(250), ...Array.from({ length: 10 }, () => u16(15)), u16(259))), /past the last column/);
  assert.throws(() => read(rec(0x00bd, u16(0), u16(255), u16(15), u32(2), u16(15), u32(2), u16(256))), /past the last column/);
  // The last row of the format is a valid row.
  assert.equal(getCell(read(number(65535, 255, 1)).sheets[0], 65535, 255)?.value, 1);
  for (const [id, name] of [[0x0203, 'NUMBER'], [0x027e, 'RK'], [0x0201, 'BLANK'], [0x00fd, 'LABELSST'], [0x0205, 'BOOLERR'], [0x0006, 'FORMULA'], [0x0208, 'ROW'], [0x007d, 'COLINFO'], [0x04bc, 'SHRFMLA'], [0x0221, 'ARRAY'], [0x0236, 'TABLE'], [0x00bd, 'MULRK'], [0x00be, 'MULBLANK']] as const) {
    assert.throws(() => read(rec(id, [1, 0, 0])), new RegExp(`a ${name} record of sheet "Sheet1" is too short`));
  }
  // A formula that states more tokens than the record holds.
  assert.throws(() => read(rec(0x0006, u16(0), u16(0), u16(15), new Array(8).fill(0), u16(0), u32(0), u16(500), [0x1e, 1, 0])), /formula of cell A1 in sheet "Sheet1" is cut off/);
  assert.throws(() => read(rec(0x04bc, u16(0), u16(1), [0, 0, 0, 0], u16(500), [1])), /SHRFMLA record .* too short/);
  // Records that only describe the view are skipped when they are too short.
  assert.doesNotThrow(() => read(cat(rec(0x023e), rec(0x0041, [1]), rec(0x001d, [1]), rec(0x00a0), rec(0x0200, [1]), rec(0x00e5), rec(0x0055), rec(0x0099), rec(0x0225), rec(0x0081), rec(0x0862, [1]))));
  // A sheet name in an error message is cut.
  assert.throws(() => readXlsStream(workbookStream({ sheets: [{ name: 'x'.repeat(200), records: number(0, 300, 1) }] })), (error: Error) => error.message.length < 200);
});

test('substreams nested inside a sheet: deep nesting, and one that never ends', () => {
  const open = Array.from({ length: 5000 }, () => bof(0x0020));
  const close = Array.from({ length: 5000 }, () => eof());
  const sheet = read(cat(number(0, 0, 1), ...open, number(9, 9, 9), ...close, number(1, 0, 2))).sheets[0];
  assert.equal(sheet.rowCount, 2);
  assert.throws(() => read(cat(number(0, 0, 1), ...open)), /is not terminated/);
  // The same in the globals.
  assert.throws(() => readXlsStream(cat(bof(0x0005), ...open, eof())), /globals are not terminated/);
});

test('shared strings: a string cut off by the end of the table, an empty table', () => {
  assert.throws(() => read(labelSst(0, 0, 0), rec(0x00fc, u32(1), u32(1), u16(50), [0], [65, 66])), /string runs past the end/);
  assert.throws(() => read(labelSst(0, 0, 0), rec(0x00fc, u32(1), u32(1), u16(2), [1], [65, 0, 66])), /split between two records|runs past the end/);
  assert.throws(() => read(labelSst(0, 0, 0)), /out of range \(0 strings\)/);
  assert.throws(() => read(labelSst(0, 0, 0), rec(0x00fc, [1, 2])), /Invalid xls/);
  // Phonetic data that claims to be longer than what is there.
  const workbook = read(labelSst(0, 0, 0), rec(0x00fc, u32(1), u32(1), u16(2), [0x04], u32(100000), [65, 66], [1, 2, 3]));
  assert.equal(getCell(workbook.sheets[0], 0, 0)?.value, 'AB');
});

test('many MULBLANK records and many column and merge records stay cheap', () => {
  const mulblank = (row: number) => rec(0x00be, u16(row), u16(0), ...Array.from({ length: 256 }, () => u16(15)), u16(255));
  const records: Uint8Array[] = [];
  for (let row = 0; row < 4000; row++) records.push(mulblank(row));
  for (let i = 0; i < 5000; i++) records.push(rec(0x007d, u16(i % 256), u16(255), u16(2000 + i), u16(15), u16(0), u16(0)));
  const merges: number[][] = [];
  for (let i = 0; i < 1000; i++) merges.push([...u16(i), ...u16(i), ...u16(0), ...u16(3)]);
  records.push(rec(0x00e5, u16(60000), ...merges));
  const started = Date.now();
  const sheet = read(cat(...records)).sheets[0];
  const elapsed = Date.now() - started;
  assert.equal(sheet.rowCount, 4000);
  assert.equal(sheet.colCount, 256);
  assert.equal(sheet.merges.length, 1000);
  assert.ok(sheet.columns.length <= 256);
  for (let i = 1; i < sheet.columns.length; i++) assert.ok(sheet.columns[i].min > sheet.columns[i - 1].max);
  assert.ok(elapsed < 5000, `took ${elapsed} ms`);
});

/** A workbook that uses most of what the reader reads, for damage tests. */
function richWorkbook(): Uint8Array {
  const tail = cat(
    supBookSelf(2),
    externSheet([0, 0, 0], [0, 1, 1]),
    definedName({ name: 'Ad', tokens: ptg.area3d(1, 0, 3, 0, 1) }),
    definedName({ name: '', builtin: 6, sheet: 1, tokens: ptg.area3d(0, 0, 3, 0, 2) }),
    sst(str16('bir'), str16('iki ğ'), cat(u16(4), [0x08], u16(1), [65, 66, 67, 68], u16(2), u16(1))),
  );
  const records = cat(
    rec(0x0200, u32(0), u32(6), u16(0), u16(4), u16(0)),
    rec(0x0208, u16(0), u16(0), u16(4), u16(300), u16(0), u16(0), u16(0x0140), u16(15)),
    rec(0x007d, u16(0), u16(1), u16(3000), u16(15), u16(2), u16(0)),
    number(0, 0, 1.5),
    labelSst(0, 1, 1),
    labelSst(0, 2, 2),
    blank(0, 3),
    formula(1, 0, tokens(ptg.ref(0, 0), ptg.int(2), ptg.mul), 3),
    formula(1, 1, tokens(ptg.name(1), ptg.funcVar(4, 1)), { text: 'metin' }),
    formula(2, 0, ptg.exp(2, 0), 1),
    shrFmla(2, 4, 0, 0, tokens(ptg.refN(0xffff, 0), ptg.int(1), ptg.add)),
    formula(3, 0, ptg.exp(2, 0), 2),
    formula(4, 0, ptg.exp(2, 0), 3),
    formula(5, 0, ptg.exp(5, 0), 4),
    arrayRec(5, 5, 0, 1, tokens(ptg.area3d(1, 0, 1, 0, 0), ptg.funcVar(4, 1))),
    rec(0x00e5, u16(1), u16(6), u16(7), u16(0), u16(2)),
    rec(0x01b8, u16(0), u16(0), u16(1), u16(1), new Array(16).fill(0), u32(2), u32(0x08), u32(3), wide('A1'), [0, 0]),
    rec(0x023e, u16(0x06b6 | 0x0008), u16(0), u16(0), u16(64), u16(0), u16(0), u16(0), u32(0)),
    rec(0x0041, u16(1), u16(1), u16(1), u16(1), [0]),
    rec(0x001d, [0], u16(2), u16(2), u16(0), u16(1), u16(2), u16(2), [2, 2]),
  );
  return workbookStream({ tail, sheets: [{ name: 'Sheet1', records }, { name: 'İkinci', records: cat(number(0, 0, 10), number(1, 0, 20)) }] });
}

test('the workbook used for the damage tests reads completely', () => {
  const { workbook, report } = readXlsStream(richWorkbook());
  const sheet = workbook.sheets[0];
  assert.equal(report.untranslatedFormulas, 0);
  assert.deepEqual([1, 2, 3, 4, 5].map((row) => getCell(sheet, row, 0)?.formula?.text), ['A1*2', 'A2+1', 'A3+1', 'A4+1', "SUM('İkinci'!$A$1:$A$2)"]);
  assert.equal(getCell(sheet, 1, 1)?.formula?.text, 'SUM(Ad)');
  assert.equal(getCell(sheet, 1, 1)?.value, 'metin');
  assert.equal(workbook.definedNames.length, 2);
  assert.equal(sheet.hyperlinks[0].location, 'A1');
  assert.deepEqual([sheet.view.frozenRows, sheet.view.frozenCols, sheet.view.activeCell], [1, 1, 'C3']);
});

test('random damage: every damaged workbook reads or fails with an Error, quickly', () => {
  const original = richWorkbook();
  // A small generator with a fixed seed, so a failure can be repeated.
  let seed = 0x2545f491;
  const random = () => {
    seed ^= seed << 13;
    seed ^= seed >>> 17;
    seed ^= seed << 5;
    return (seed >>> 0) / 0x100000000;
  };
  const started = Date.now();
  let read = 0;
  let failed = 0;
  for (let round = 0; round < 3000; round++) {
    const copy = original.slice();
    const changes = 1 + Math.floor(random() * 4);
    for (let i = 0; i < changes; i++) {
      const at = Math.floor(random() * copy.length);
      const kind = random();
      copy[at] = kind < 0.3 ? 0xff : kind < 0.6 ? 0 : Math.floor(random() * 256);
    }
    try {
      const { workbook } = readXlsStream(copy);
      assert.ok(workbook.sheets.length > 0);
      read++;
    } catch (error) {
      assert.ok(error instanceof Error, `round ${round}: ${String(error)}`);
      assert.ok(!(error instanceof TypeError) && !(error instanceof RangeError), `round ${round}: ${(error as Error).stack}`);
      failed++;
    }
  }
  const elapsed = Date.now() - started;
  assert.equal(read + failed, 3000);
  assert.ok(read > 0 && failed > 0);
  assert.ok(elapsed < 20000, `took ${elapsed} ms`);
});

test('performance: a sheet of 65,536 rows and 20 columns', () => {
  const rows = 65536;
  const cols = 20;
  const strings = 500;
  // Built straight into one buffer: 16 number cells (NUMBER and RK), 2 shared strings and 2 formulas per row.
  const head = cat(workbookStream({ sheets: [{ name: 'Büyük' }], tail: sst(...Array.from({ length: strings }, (_, i) => str16(`Metin ${i} ğüşiöç`))) }));
  // The empty sheet is the last 20 + 4 bytes: BOF (20) and EOF (4). The cells go between them.
  const sheetStart = head.length - 24;
  const perRow = 8 * 18 + 8 * 14 + 2 * 14 + 2 * (26 + 11);
  const out = new Uint8Array(head.length + rows * perRow);
  out.set(head.subarray(0, sheetStart + 20));
  const view = new DataView(out.buffer);
  let p = sheetStart + 20;
  const header = (id: number, size: number, row: number, col: number) => {
    view.setUint16(p, id, true);
    view.setUint16(p + 2, size, true);
    view.setUint16(p + 4, row, true);
    view.setUint16(p + 6, col, true);
    view.setUint16(p + 8, 15, true);
    p += 10;
  };
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < 8; col++) {
      header(0x0203, 14, row, col);
      view.setFloat64(p, row * 1.25 + col, true);
      p += 8;
    }
    for (let col = 8; col < 16; col++) {
      header(0x027e, 10, row, col);
      view.setUint32(p, ((((row + col) % 100000) << 2) | 2) >>> 0, true);
      p += 4;
    }
    for (let col = 16; col < 18; col++) {
      header(0x00fd, 10, row, col);
      view.setUint32(p, (row + col) % strings, true);
      p += 4;
    }
    for (let col = 18; col < 20; col++) {
      // =A{row}+I{row}
      header(0x0006, 22 + 11, row, col);
      view.setFloat64(p, row * 2.25 + 8, true);
      view.setUint16(p + 14, 11, true);
      p += 16;
      out.set([0x44, row & 0xff, row >> 8, 0x00, 0xc0, 0x44, row & 0xff, row >> 8, 0x08, 0xc0, 0x03], p);
      p += 11;
    }
  }
  out.set(eof(), p);
  assert.equal(p + 4, out.length);

  const started = Date.now();
  const { workbook, report } = readXlsStream(out);
  const elapsed = Date.now() - started;
  const sheet = workbook.sheets[0];
  assert.equal(sheet.rowCount, rows);
  assert.equal(sheet.colCount, cols);
  assert.equal(sheet.rows.size, rows);
  assert.equal(getCell(sheet, 65535, 7)?.value, 65535 * 1.25 + 7);
  assert.equal(getCell(sheet, 12345, 15)?.value, (12345 + 15) % 100000);
  assert.equal(getCell(sheet, 500, 17)?.value, 'Metin 17 ğüşiöç');
  assert.equal(getCell(sheet, 65535, 19)?.formula?.text, 'A65536+I65536');
  assert.equal(report.untranslatedFormulas, 0);
  // Measured at about 0.4 s on a laptop; the limit leaves room for a slow machine.
  assert.ok(elapsed < 8000, `reading took ${elapsed} ms`);
  console.log(`xls 65,536 x 20 cells: ${elapsed} ms`);
});

test('strings in records the reader takes names from are bounded by the record', () => {
  // A sheet name and a font name that claim more characters than the record holds.
  assert.throws(() => readXlsStream(cat(bof(0x0005), rec(0x0085, u32(0), [0, 0], [200, 0, 65]), eof())), /string runs past the end/);
  assert.throws(() => readXlsStream(cat(bof(0x0005), rec(0x0031, u16(200), u16(0), u16(0), u16(400), u16(0), [0, 0, 0, 0], [250, 1, 65, 0]), eof())), /Invalid xls/);
  assert.throws(() => readXlsStream(cat(bof(0x0005), rec(0x0018, u16(0), [0, 200], u16(0), u16(0), u16(0), [0, 0, 0, 0], [0], [65]), eof())), /Invalid xls/);
  assert.ok(str8);
});
