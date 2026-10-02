// Limits of the xls reader: what a small file may ask for in memory and time,
// and the corners of the BIFF5 layout.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { MAX_FORMULA_CHARS, MAX_FORMULA_DEPTH, formulaText, type FormulaContext } from '../src/formats/xls/formula';
import { readerAt, recordStream } from '../src/formats/xls/records';
import { XLS_MAX_SHEETS, readXlsStream } from '../src/formats/xls/workbook';
import { resolveCellStyle } from '../src/model/sheet/color';
import { getCell } from '../src/model/sheet/types';
import { explicitFormula, sharedMasters } from '../src/ops/ops';
import { cat, defaultStyles, eof, externSheet, formula, narrow, number, ptg, rec, shrFmla, supBookSelf, tokens, u16, u32, wide, workbookStream } from './xls-fixtures';

const context = (): FormulaContext => ({ biff8: true, codepage: 1252, sheetNames: ['Sheet1'], names: [], xti: [], supBooks: [], failures: new Map() });

function text(parts: number[], extra: number[] = [], ctx = context()): string | undefined {
  const bytes = new Uint8Array([...parts, ...extra]);
  return formulaText(ctx, bytes, 0, parts.length, parts.length, bytes.length, 0, 0);
}

const repeat = (part: number[], count: number): number[] => {
  const out: number[] = [];
  for (let i = 0; i < count; i++) for (const byte of part) out.push(byte);
  return out;
};

test('formula text: parentheses and function calls nest as deep as Excel allows and no deeper', () => {
  assert.equal(MAX_FORMULA_DEPTH, 64);
  const ctx = context();
  const deepest = text([...ptg.int(1), ...repeat(ptg.paren, 64)], [], ctx);
  assert.equal(deepest, '('.repeat(64) + '1' + ')'.repeat(64));
  assert.equal(text([...ptg.int(1), ...repeat(ptg.paren, 65)], [], ctx), undefined);
  assert.equal(ctx.failures.get('formula nested too deeply'), 1);

  // ABS(ABS(...ABS(1))): function calls count as parentheses do, and so do the two mixed.
  const abs = ptg.func(24);
  assert.equal(text([...ptg.int(1), ...repeat(abs, 64)]), 'ABS('.repeat(64) + '1' + ')'.repeat(64));
  assert.equal(text([...ptg.int(1), ...repeat(abs, 65)]), undefined);
  assert.equal(text([...ptg.int(1), ...repeat([...abs, ...ptg.paren], 33)]), undefined);
  assert.equal(text([...ptg.int(1), ...repeat(ptg.funcVar(4, 1), 65)]), undefined);
  assert.equal(text([...ptg.int(1), ...repeat(ptg.attrSum, 65)]), undefined);
  // Nesting is that of the deepest operand, not the sum over the operands.
  const side = [...ptg.int(1), ...repeat(ptg.paren, 40)];
  assert.equal(text([...side, ...side, ...ptg.add])?.length, 2 * 81 + 1);
});

test('formula text: longer than 8,192 characters is not a formula', () => {
  assert.equal(MAX_FORMULA_CHARS, 8192);
  const word = 'x'.repeat(250);
  // 252 characters a string, one more for each "&".
  const joined = (count: number) => [...ptg.str(word), ...repeat([...ptg.str(word), ...ptg.concat], count - 1)];
  const ctx = context();
  assert.equal(text(joined(32), [], ctx)?.length, 32 * 253 - 1);
  assert.equal(text(joined(33), [], ctx), undefined);
  assert.equal(ctx.failures.get('formula text too long'), 1);
  // Operators that add a character at a time.
  assert.equal(text([...ptg.int(1), ...repeat(ptg.percent, 8191)])?.length, 8192);
  assert.equal(text([...ptg.int(1), ...repeat(ptg.percent, 8192)]), undefined);
  assert.equal(text([...ptg.int(1), ...repeat(ptg.uminus, 8192)]), undefined);
  // An array constant of 256 x 40 numbers.
  const cells: number[] = [];
  for (let i = 0; i < 256 * 40; i++) cells.push(0x01, 0, 0, 0, 0, 0, 0, 0xf0, 0x3f);
  assert.equal(text(ptg.array(), [255, ...u16(39), ...cells]), undefined);
  assert.equal(text(ptg.array(), [255, ...u16(0), ...cells.slice(0, 256 * 9)])?.length, 256 * 2 + 1);
  // A string of an array constant that is longer than a formula.
  assert.equal(text(ptg.array(), [0, ...u16(0), 0x02, ...u16(9000), 0, ...new Array(9000).fill(0x41)]), undefined);
});

/** A sheet whose column A is one shared formula of `count` cells with the given tokens. */
function sharedSheet(count: number, shared: number[]): Uint8Array {
  const parts: Uint8Array[] = [];
  for (let row = 0; row < count; row++) {
    parts.push(formula(row, 0, ptg.exp(0, 0), row + 0.5));
    if (row === 0) parts.push(shrFmla(0, count - 1, 0, 0, shared));
  }
  return workbookStream({ sheets: [{ name: 'Sheet1', records: cat(...parts) }] });
}

test('a shared formula of thousands of nested parentheses costs nothing: the cells keep their values', () => {
  const count = 6000;
  const stream = sharedSheet(count, [...ptg.int(1), ...repeat(ptg.paren, 8000)]);
  const before = process.memoryUsage().heapUsed;
  const started = Date.now();
  const { workbook, report } = readXlsStream(stream);
  const elapsed = Date.now() - started;
  const sheet = workbook.sheets[0];
  // Every cell is counted: as too deeply nested, or as left out once the group has cost too much.
  assert.equal(report.untranslatedFormulas, count);
  const nested = report.untranslatedReasons.get('formula nested too deeply') || 0;
  const leftOut = report.untranslatedReasons.get('shared formula: the texts of the workbook together are too large') || 0;
  assert.ok(nested > 0);
  assert.equal(nested + leftOut, count);
  for (const row of [0, 1, count - 1]) assert.deepEqual(getCell(sheet, row, 0), { kind: 'number', value: row + 0.5, style: 0 });
  assert.ok(elapsed < 3000, `took ${elapsed} ms`);
  // The texts this once made took more than a gigabyte.
  const grown = process.memoryUsage().heapUsed - before;
  assert.ok(grown < 128 * 1024 * 1024, `the heap grew by ${Math.round(grown / 1048576)} MB`);
});

test('shared formulas: one text for a group however many cells it has, and every cell has its formula', () => {
  const word = 'y'.repeat(250);
  const shared = [...ptg.str(word), ...repeat([...ptg.str(word), ...ptg.concat], 29), ...ptg.refN(0, 1), ...ptg.concat];
  const count = 2500;
  const stream = sharedSheet(count, shared);
  const { workbook, report } = readXlsStream(stream);
  const sheet = workbook.sheets[0];
  const masters = sharedMasters(sheet);
  let chars = 0;
  for (let row = 0; row < count; row++) {
    const cell = getCell(sheet, row, 0)!;
    assert.equal(cell.value, row + 0.5);
    assert.ok(cell.formula, `row ${row} has no formula`);
    chars += cell.formula.text.length;
  }
  // The first cell holds the text; the others point to it, checked or (once reading the group has cost too much) not.
  assert.equal(getCell(sheet, 0, 0)?.formula?.ref, `A1:A${count}`);
  assert.equal(chars, 30 * 253 - 1 + 2 + 1);
  for (const row of [1, 100, count - 1]) {
    const cell = getCell(sheet, row, 0)!;
    assert.deepEqual(cell.formula, { text: '', type: 'shared', si: 0 });
    const text = explicitFormula(cell, row, 0, masters).formula!.text;
    assert.equal(text.length, 30 * 253 - 1 + 2 + String(row + 1).length);
    assert.ok(text.endsWith('&B' + (row + 1)));
  }
  assert.equal(report.untranslatedFormulas, 0);

  // A genuine group is far below the bound: every cell is checked and gets its formula.
  const plain = readXlsStream(sharedSheet(3000, tokens(ptg.refN(0, 1), ptg.int(2), ptg.mul)));
  assert.equal(plain.report.untranslatedFormulas, 0);
  const last = getCell(plain.workbook.sheets[0], 2999, 0)!;
  assert.equal(explicitFormula(last, 2999, 0, sharedMasters(plain.workbook.sheets[0])).formula?.text, 'B3000*2');
});

const STD_LINK = [0xd0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b];
const FILE_MONIKER = [0x03, 0x03, 0, 0, 0, 0, 0, 0, 0xc0, 0, 0, 0, 0, 0, 0, 0x46];
const linkString = (value: string) => [...u32(value.length + 1), ...wide(value), 0, 0];
const hlink = (top: number, bottom: number, left: number, right: number, flags: number, ...body: number[][]) => rec(0x01b8, u16(top), u16(bottom), u16(left), u16(right), STD_LINK, u32(2), u32(flags), ...body);
const tooltip = (top: number, bottom: number, left: number, right: number, value: string) => rec(0x0800, u16(0x0800), u16(top), u16(bottom), u16(left), u16(right), wide(value), [0, 0]);

test('tooltips find their links by range: the last one for a range wins, an empty one changes nothing', () => {
  const sheet = readXlsStream(
    workbookStream({
      sheets: [
        {
          name: 'Sheet1',
          records: cat(
            hlink(0, 0, 0, 0, 0x08, linkString('A1')),
            hlink(1, 2, 3, 4, 0x08, linkString('B1')),
            // Two links over the same range share the tooltip.
            hlink(1, 2, 3, 4, 0x08, linkString('C1')),
            hlink(5, 5, 0, 0, 0x08, linkString('D1')),
            tooltip(1, 2, 3, 4, 'first'),
            tooltip(1, 2, 3, 4, 'second'),
            tooltip(1, 2, 3, 4, ''),
            tooltip(0, 0, 0, 0, 'top'),
            // The same corners in another order are another range.
            tooltip(5, 0, 5, 0, 'none'),
            tooltip(9, 9, 9, 9, 'no link'),
          ),
        },
      ],
    }),
  ).workbook.sheets[0];
  assert.deepEqual(sheet.hyperlinks.map((link) => [link.location, link.tooltip]), [['A1', 'top'], ['B1', 'second'], ['C1', 'second'], ['D1', undefined]]);
});

test('tens of thousands of links and tooltips are matched in one pass', () => {
  const parts: Uint8Array[] = [];
  const count = 40000;
  for (let i = 0; i < count; i++) parts.push(hlink(i, i, 0, 0, 0x08, linkString('A1')));
  for (let i = 0; i < count; i++) parts.push(tooltip(i, i, 1, 1, 't'));
  parts.push(tooltip(count - 1, count - 1, 0, 0, 'last'));
  const stream = workbookStream({ sheets: [{ name: 'Sheet1', records: cat(...parts) }] });
  const started = Date.now();
  const sheet = readXlsStream(stream).workbook.sheets[0];
  const elapsed = Date.now() - started;
  assert.equal(sheet.hyperlinks.length, count);
  assert.equal(sheet.hyperlinks[count - 1].tooltip, 'last');
  assert.equal(sheet.hyperlinks[0].tooltip, undefined);
  // Comparing every tooltip with every link took seconds.
  assert.ok(elapsed < 1500, `took ${elapsed} ms`);
});

test('a record continued in a hundred thousand pieces: what is left is known without counting again', () => {
  // EXTERNSHEET with 20,000 entries, each of its bytes in a CONTINUE record of its own.
  const entries = 20000;
  const body: number[] = [];
  for (let i = 0; i < entries; i++) body.push(0, 0, 0, 0, 0, 0);
  const first = cat(supBookSelf(1), rec(0x0017, u16(entries)));
  const tail = new Uint8Array(first.length + body.length * 5);
  tail.set(first);
  for (let i = 0, at = first.length; i < body.length; i++, at += 5) tail.set([0x3c, 0, 1, 0, body[i]], at);
  const stream = workbookStream({ tail, sheets: [{ name: 'Sheet1', records: formula(0, 0, ptg.ref3d(entries - 1, 4, 2), 1) }] });
  const started = Date.now();
  const { workbook } = readXlsStream(stream);
  const elapsed = Date.now() - started;
  assert.equal(getCell(workbook.sheets[0], 0, 0)?.formula?.text, 'Sheet1!$C$5');
  // Adding up the pieces again for every entry took over a second.
  assert.ok(elapsed < 500, `took ${elapsed} ms`);

  // The count itself: bytes left over all pieces, and "done" only when every piece is used up.
  const bytes = cat(rec(0x0017, [1, 2]), rec(0x003c, []), rec(0x003c, [3]), rec(0x003c, [4, 5, 6]), rec(0x003c, []), eof());
  const r = readerAt(recordStream(bytes), 0);
  assert.equal(r.remaining, 6);
  assert.equal(r.u16(), 0x0201);
  assert.equal(r.remaining, 4);
  assert.equal(r.done, false);
  assert.equal(r.u8(), 3);
  assert.equal(r.remaining, 3);
  assert.equal(r.u16(), 0x0504);
  assert.equal(r.remaining, 1);
  assert.equal(r.done, false);
  r.skip(1);
  assert.equal(r.remaining, 0);
  assert.equal(r.done, true);
  assert.equal(r.rest().length, 0);
  assert.throws(() => r.u8(), /Invalid xls/);
});

test('HLINK: the path of a file moniker is in the code page of the workbook', () => {
  const path = [0xfe, 0x75, 0x62, 0x65, 0x5c, 0xfd, 0x2e, 0x78, 0x6c, 0x73, 0]; // "şube\ı.xls" in Windows-1254
  const link = hlink(0, 0, 0, 0, 0x01, FILE_MONIKER, [...u16(0), ...u32(path.length), ...path], new Array(24).fill(0), u32(0));
  const read = (codepage: number | undefined) =>
    readXlsStream(workbookStream({ head: cat(codepage === undefined ? [] : rec(0x0042, u16(codepage)), defaultStyles()), sheets: [{ name: 'Sheet1', records: link }] })).workbook.sheets[0].hyperlinks;
  assert.deepEqual(read(1254), [{ range: { top: 0, left: 0, bottom: 0, right: 0 }, target: 'şube/ı.xls' }]);
  assert.equal(read(1252)[0].target, 'þube/ý.xls');
  assert.equal(read(undefined)[0].target, 'þube/ý.xls');
});

test('a workbook of nothing but sheets is refused past the limit, with UnsupportedFileError', () => {
  const sheets = (count: number) => Array.from({ length: count }, (_, i) => ({ name: 'S' + i }));
  assert.equal(XLS_MAX_SHEETS, 4096);
  assert.equal(readXlsStream(workbookStream({ sheets: sheets(XLS_MAX_SHEETS) })).workbook.sheets.length, XLS_MAX_SHEETS);
  assert.throws(
    () => readXlsStream(workbookStream({ sheets: sheets(XLS_MAX_SHEETS + 1) })),
    (error: Error) => error.name === 'UnsupportedFileError' && /more than 4096 sheets/.test(error.message),
  );
});

test('MULRK: the last two bytes are the last column, never a cell', () => {
  const sheet = (record: Uint8Array) => readXlsStream(workbookStream({ sheets: [{ name: 'Sheet1', records: record }] })).workbook.sheets[0];
  const rkOf = (value: number) => u32(((value << 2) | 2) >>> 0);
  // Two cells and the last column.
  const whole = sheet(rec(0x00bd, u16(0), u16(3), u16(15), rkOf(7), u16(15), rkOf(8), u16(4)));
  assert.deepEqual([2, 3, 4, 5].map((col) => getCell(whole, 0, col)?.value), [undefined, 7, 8, undefined]);
  assert.equal(whole.colCount, 5);
  // A record cut off after its cells: the bytes where the last column would be are not read as the end of a cell.
  const cut = sheet(rec(0x00bd, u16(0), u16(3), u16(15), rkOf(7), u16(15), rkOf(8)));
  assert.deepEqual([3, 4].map((col) => getCell(cut, 0, col)?.value), [7, undefined]);
  // One cell, the last column, and nothing else.
  const single = sheet(rec(0x00bd, u16(2), u16(255), u16(15), rkOf(9), u16(255)));
  assert.equal(getCell(single, 2, 255)?.value, 9);
  assert.equal(single.colCount, 256);
});

// ---------------------------------------------------------------------------
// BIFF5

/** A BIFF5 FONT record: the name is a byte string without option flags. */
const font5 = (name: number[], weight = 400) => rec(0x0031, u16(200), u16(0), u16(0x7fff), u16(weight), u16(0), [0, 0, 162, 0], [name.length], name);

/** A BIFF5 XF record: 16 bytes. */
const xf5 = (fontIndex: number, format: number, type: number, align: number, orientation: number, fill: number, border: number) =>
  rec(0x00e0, u16(fontIndex), u16(format), u16(type), [align, orientation], u32(fill), u32(border));

/** A BIFF5 NAME record: the name is a byte string, the tokens follow it. */
const name5 = (name: number[], nameTokens: number[], sheet = 0) => rec(0x0018, u16(0), [0, name.length], u16(nameTokens.length), u16(0), u16(sheet), [0, 0, 0, 0], name, nameTokens);

/** The 14 bytes in front of a BIFF5 3D reference to sheets of this workbook. */
const tabs5 = (first: number, last: number) => [...u16(0xffff), 0, 0, 0, 0, 0, 0, 0, 0, ...u16(first), ...u16(last)];

test('BIFF5: XF records of 16 bytes, EXTERNSHEET and NAME records, references with 14 bits for the row', () => {
  const styleXf = xf5(0, 0, 0xfff5, 0x20, 0, 0x20c0, 0);
  const fancy = xf5(
    1,
    2,
    0x0001,
    0x02 | 0x08 | (1 << 4),
    2,
    (10 | (11 << 7) | (2 << 16) | (5 << 22) | (12 << 25)) >>> 0,
    (1 | (2 << 3) | (3 << 6) | (13 << 9) | (14 << 16) | (15 << 23)) >>> 0,
  );
  const head = cat(
    rec(0x0042, u16(1254)),
    font5(narrow('Arial')),
    font5([0xdd, 0x6e, 0x63, 0x65], 700), // "İnce"
    ...Array.from({ length: 15 }, () => styleXf),
    xf5(0, 0, 0x0001, 0x20, 0, 0x20c0, 0),
    fancy,
  );
  const tail = cat(
    // A sheet of this workbook by name, then another workbook.
    rec(0x0017, [3, 0x03, 0x4b, 0xfd, 0xfe]), // "Kış"
    rec(0x0017, [5, 0x01, ...narrow('a.xls')]),
    name5([0xde, 0x61, 0x72, 0x74], [0x3a, ...u16(1), 0, 0, 0, 0, 0, 0, 0, 0, ...u16(0), ...u16(0), ...u16(2), 3]), // "Şart" -> the sheet of the first EXTERNSHEET
    name5(narrow('Yerel'), [0x3b, ...tabs5(0, 1), ...u16(0), ...u16(0x3fff), 1, 2], 2),
    name5(narrow('Dis'), [0x3a, ...u16(2), 0, 0, 0, 0, 0, 0, 0, 0, ...u16(0), ...u16(0), ...u16(0), 0]),
  );
  // The shared formula: one row up (wrapping inside 14 bits), same column, times the first name.
  const shared = [0x4c, ...u16(0x3fff | 0xc000), 0, 0x43, ...u16(1), ...new Array(12).fill(0), 0x05];
  const cell = (row: number, col: number, cellTokens: number[], value: number) => rec(0x0006, u16(row), u16(col), u16(16), [...new Uint8Array(new Float64Array([value]).buffer)], u16(0), u32(0), u16(cellTokens.length), cellTokens);
  const records = cat(
    rec(0x0200, u16(0), u16(16384), u16(0), u16(3), u16(0)),
    rec(0x0204, u16(0), u16(0), u16(16), u16(3), [0xde, 0x75, 0xfe]), // "Şuş"
    cell(0, 1, [0x01, ...u16(0), ...u16(1)], 1),
    rec(0x04bc, u16(0), u16(1), [1, 1, 0, 0], u16(shared.length), shared),
    cell(1, 1, [0x01, ...u16(0), ...u16(1)], 2),
    // The last row a BIFF5 sheet has, named with all 14 bits; an area over all rows is a whole column.
    cell(16383, 2, [0x24, ...u16(0x3fff | 0xc000), 2, 0x25, ...u16(0), ...u16(0x3fff), 0, 0, 0x42, 1, ...u16(4), 0x03], 3),
  );
  const { workbook, report } = readXlsStream(
    // The sheet name as the bytes of the code page: "Kış".
    workbookStream({ version: 0x0500, head, tail, sheets: [{ name: String.fromCharCode(0x4b, 0xfd, 0xfe), records }, { name: 'B', records: number(0, 0, 1, 15) }] }),
    'Book',
  );
  assert.equal(report.biff, 5);
  assert.equal(report.codepage, 1254);
  assert.equal(report.untranslatedFormulas, 0);
  const sheet = workbook.sheets[0];
  assert.deepEqual(sheet.dimension, { top: 0, left: 0, bottom: 16383, right: 2 });
  assert.equal(getCell(sheet, 0, 0)?.value, 'Şuş');
  assert.equal(getCell(sheet, 0, 1)?.formula?.text, 'B16384*Şart');
  assert.equal(getCell(sheet, 1, 1)?.formula?.text, 'B1*Şart');
  assert.equal(getCell(sheet, 16383, 2)?.formula?.text, 'C16384+SUM($A:$A)');
  assert.equal(sheet.rowCount, 16384);

  assert.deepEqual(workbook.definedNames, [
    { name: 'Şart', value: "'Kış'!$D$3", localSheet: undefined, hidden: false },
    { name: 'Yerel', value: "'Kış:B'!$B:$C", localSheet: 1, hidden: false },
  ]);
  // A reference to another workbook is not written as text in BIFF5.
  assert.equal(report.untranslatedNames, 1);

  const style = resolveCellStyle(workbook.styles, getCell(sheet, 0, 0)?.style ?? -1);
  assert.equal(style.font.name, 'İnce');
  assert.equal(style.font.bold, true);
  assert.equal(style.numFmt, '0.00');
  assert.deepEqual(style.alignment, { horizontal: 'center', vertical: 'center', wrapText: true, textRotation: 90 });
  assert.deepEqual(style.fill, { type: 'pattern', pattern: 'mediumGray', fg: { indexed: 10 }, bg: { indexed: 11 } });
  assert.deepEqual(style.border, {
    diagonalUp: false,
    diagonalDown: false,
    left: { style: 'medium', color: { indexed: 14 } },
    right: { style: 'dashed', color: { indexed: 15 } },
    top: { style: 'thin', color: { indexed: 13 } },
    bottom: { style: 'thick', color: { indexed: 12 } },
  });
  // The cell records after the XF records were read from the right place.
  assert.equal(getCell(workbook.sheets[1], 0, 0)?.style, 0);
});

test('BIFF5: a sheet named by an EXTERNSHEET record that formulas of another sheet refer to', () => {
  const head = cat(rec(0x0042, u16(1254)), font5(narrow('Arial')), ...Array.from({ length: 16 }, () => xf5(0, 0, 0x0001, 0x20, 0, 0x20c0, 0)));
  const tail = rec(0x0017, [2, 0x03, 0xd6, 0x7a]); // "Öz"
  const cell = rec(0x0006, u16(0), u16(0), u16(15), [0, 0, 0, 0, 0, 0, 0x14, 0x40], u16(0), u32(0), u16(18), [0x3a, ...u16(1), 0, 0, 0, 0, 0, 0, 0, 0, ...u16(0), ...u16(0), ...u16(1 | 0xc000), 1]);
  const { workbook } = readXlsStream(workbookStream({ version: 0x0500, head, tail, sheets: [{ name: 'A', records: cell }, { name: String.fromCharCode(0xd6, 0x7a), records: number(1, 1, 5, 15) }] }), 'Book');
  assert.equal(workbook.sheets[1].name, 'Öz');
  assert.equal(getCell(workbook.sheets[0], 0, 0)?.value, 5);
  assert.deepEqual(getCell(workbook.sheets[0], 0, 0)?.formula, { text: "'Öz'!B2", type: 'normal' });
});

test('EXTERNSHEET (BIFF8): entries are read to the end of the record and not beyond', () => {
  // The count says three entries, the record holds two: the third is not made up.
  const tail = cat(supBookSelf(1), rec(0x0017, u16(3), u16(0), u16(0), u16(0), u16(0), u16(0), u16(0)));
  const records = cat(formula(0, 0, ptg.ref3d(1, 0, 0), 1), formula(1, 0, ptg.ref3d(2, 0, 0), 1));
  const { workbook, report } = readXlsStream(workbookStream({ tail, sheets: [{ name: 'Sheet1', records }] }));
  assert.equal(getCell(workbook.sheets[0], 0, 0)?.formula?.text, 'Sheet1!$A$1');
  assert.equal(getCell(workbook.sheets[0], 1, 0)?.formula, undefined);
  assert.equal(report.untranslatedFormulas, 1);
  assert.deepEqual(externSheet([0, 0, 0]).length, 4 + 2 + 6);
});
