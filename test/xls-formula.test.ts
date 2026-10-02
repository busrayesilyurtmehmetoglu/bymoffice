import assert from 'node:assert/strict';
import { test } from 'node:test';
import { parseFormula } from '../src/calc/parse';
import { CalcEngine } from '../src/calc/engine';
import { formulaText, quoteSheetName, type FormulaContext } from '../src/formats/xls/formula';
import { fixedArgCount, functionName } from '../src/formats/xls/functions';
import { readXlsStream } from '../src/formats/xls/workbook';
import { getCell } from '../src/model/sheet/types';
import { explicitFormula, sharedMasters } from '../src/ops/ops';
import {
  arrayRec,
  cat,
  definedName,
  externName,
  externSheet,
  f64,
  formula,
  narrow,
  number,
  ptg,
  rec,
  shrFmla,
  str16,
  supBookAddin,
  supBookExternal,
  supBookSelf,
  tokens,
  u16,
  workbookStream,
} from './xls-fixtures';

function context(overrides: Partial<FormulaContext> = {}): FormulaContext {
  return {
    biff8: true,
    codepage: 1252,
    sheetNames: ['Sheet1', 'Veri sayfası', "Ali'nin", 'Q1'],
    names: ['Toplam', 'KDV_Oranı', '_xlfn.IFERROR'],
    xti: [
      { book: 0, first: 0, last: 0 },
      { book: 0, first: 1, last: 1 },
      { book: 0, first: 0, last: 3 },
      { book: 0, first: -1, last: -1 },
      { book: 1, first: 0, last: 0 },
      { book: 1, first: 1, last: 1 },
      { book: 2, first: -2, last: -2 },
      { book: 0, first: -2, last: -2 },
      { book: 1, first: -2, last: -2 },
      { book: 0, first: 2, last: 2 },
      { book: 0, first: 3, last: 3 },
    ],
    supBooks: [
      { kind: 'self', sheetNames: [], externNames: [], ordinal: 0 },
      { kind: 'external', sheetNames: ['Data', 'Dış sayfa'], externNames: ['Kur', ''], ordinal: 1 },
      { kind: 'addin', sheetNames: [], externNames: ['EOMONTH', 'XIRR'], ordinal: 0 },
    ],
    failures: new Map(),
    ...overrides,
  };
}

/** Formula text for tokens, read for the cell C5 (row 4, column 2). */
function text(parts: number[], extra: number[] = [], ctx = context()): string | undefined {
  const bytes = new Uint8Array([...parts, ...extra]);
  return formulaText(ctx, bytes, 0, parts.length, parts.length, bytes.length, 4, 2);
}

test('operands: integers, numbers, strings, booleans, errors, missing arguments', () => {
  assert.equal(text(ptg.int(65535)), '65535');
  assert.equal(text(ptg.num(0.1)), '0.1');
  assert.equal(text(ptg.num(1e21)), '1E+21');
  assert.equal(text(ptg.num(2.5e-7)), '2.5E-7');
  assert.equal(text(ptg.num(-3)), '-3');
  assert.equal(text(ptg.str('a "b" ç')), '"a ""b"" ç"');
  assert.equal(text([0x17, 2, 1, 0x5e, 0x01, 0x5f, 0x01]), '"Şş"');
  assert.equal(text(ptg.str('')), '""');
  assert.equal(text(ptg.bool(true)), 'TRUE');
  assert.equal(text(ptg.bool(false)), 'FALSE');
  assert.deepEqual([0x00, 0x07, 0x0f, 0x17, 0x1d, 0x24, 0x2a].map((code) => text(ptg.err(code))), ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A']);
  assert.equal(text(tokens(ptg.int(1), ptg.missArg, ptg.int(3), ptg.funcVar(1, 3))), 'IF(1,,3)');
});

test('every operator, in the order of the tokens', () => {
  const pair = (op: number[]) => text(tokens(ptg.ref(0, 0), ptg.ref(1, 1), op));
  assert.deepEqual(
    [ptg.add, ptg.sub, ptg.mul, ptg.div, ptg.power, ptg.concat, ptg.lt, ptg.le, ptg.eq, ptg.ge, ptg.gt, ptg.ne, ptg.isect, ptg.union, ptg.range].map(pair),
    ['A1+B2', 'A1-B2', 'A1*B2', 'A1/B2', 'A1^B2', 'A1&B2', 'A1<B2', 'A1<=B2', 'A1=B2', 'A1>=B2', 'A1>B2', 'A1<>B2', 'A1 B2', 'A1,B2', 'A1:B2'],
  );
  assert.equal(text(tokens(ptg.ref(0, 0), ptg.uminus)), '-A1');
  assert.equal(text(tokens(ptg.ref(0, 0), ptg.uplus)), '+A1');
  assert.equal(text(tokens(ptg.int(5), ptg.percent)), '5%');
  assert.equal(text(tokens(ptg.int(1), ptg.int(2), ptg.add, ptg.paren, ptg.int(3), ptg.mul)), '(1+2)*3');
});

test('parentheses are added where the tokens need them and the file has none', () => {
  // (1+2)*3 without a PtgParen.
  assert.equal(text(tokens(ptg.int(1), ptg.int(2), ptg.add, ptg.int(3), ptg.mul)), '(1+2)*3');
  // 1-(2-3): operators group from the left.
  assert.equal(text(tokens(ptg.int(1), ptg.int(2), ptg.int(3), ptg.sub, ptg.sub)), '1-(2-3)');
  assert.equal(text(tokens(ptg.int(1), ptg.int(2), ptg.sub, ptg.int(3), ptg.sub)), '1-2-3');
  // 2^(3^2), and -(2^2), which is not what -2^2 means in Excel.
  assert.equal(text(tokens(ptg.int(2), ptg.int(3), ptg.int(2), ptg.power, ptg.power)), '2^(3^2)');
  assert.equal(text(tokens(ptg.int(2), ptg.int(2), ptg.power, ptg.uminus)), '-(2^2)');
  assert.equal(text(tokens(ptg.int(2), ptg.uminus, ptg.int(2), ptg.power)), '-2^2');
  assert.equal(text(tokens(ptg.int(1), ptg.int(2), ptg.add, ptg.percent)), '(1+2)%');
  assert.equal(text(tokens(ptg.int(1), ptg.int(2), ptg.concat, ptg.int(3), ptg.eq)), '1&2=3');
  assert.equal(text(tokens(ptg.int(1), ptg.int(2), ptg.int(3), ptg.eq, ptg.concat)), '1&(2=3)');
  // A union as a function argument would read as two arguments.
  assert.equal(text(tokens(ptg.ref(0, 0), ptg.ref(1, 1), ptg.union, ptg.funcVar(4, 1))), 'SUM((A1,B2))');
  assert.equal(text(tokens(ptg.num(-2), ptg.int(2), ptg.power)), '-2^2');
});

test('references: relative and absolute, areas, whole columns and rows', () => {
  assert.equal(text(ptg.ref(9, 27)), 'AB10');
  assert.equal(text(ptg.ref(9, 27, false, false)), '$AB$10');
  assert.equal(text(ptg.ref(9, 27, true, false)), '$AB10');
  assert.equal(text(ptg.ref(9, 27, false, true)), 'AB$10');
  assert.equal(text(ptg.ref(65535, 255)), 'IV65536');
  assert.equal(text(ptg.area(0, 9, 0, 2)), 'A1:C10');
  assert.equal(text(ptg.area(0, 9, 0, 2, false)), '$A$1:$C$10');
  assert.equal(text(ptg.area(0, 65535, 1, 2, false)), '$B:$C');
  assert.equal(text(ptg.area(0, 65535, 1, 1)), 'B:B');
  assert.equal(text(ptg.area(2, 4, 0, 255, false)), '$3:$5');
  assert.equal(text(ptg.area(0, 65535, 0, 255)), 'A1:IV65536');
  // Value and array classes of the same tokens.
  assert.equal(text(ptg.ref(0, 0, true, true, 0x24)), 'A1');
  assert.equal(text(ptg.ref(0, 0, true, true, 0x64)), 'A1');
  assert.equal(text(ptg.area(0, 1, 0, 1, true, 0x45)), 'A1:B2');
  assert.equal(text(ptg.refErr), '#REF!');
  assert.equal(text(tokens(ptg.areaErr, ptg.funcVar(4, 1))), 'SUM(#REF!)');
});

test('relative tokens (PtgRefN, PtgAreaN) count from the cell of the formula and wrap around', () => {
  // Read for C5.
  assert.equal(text(ptg.refN(0xffff, 0xff)), 'B4');
  assert.equal(text(ptg.refN(1, 2)), 'E6');
  assert.equal(text(ptg.refN(7, 3, false, false)), '$D$8');
  assert.equal(text(ptg.refN(7, 0xfe, false, true)), 'A$8');
  assert.equal(text(ptg.areaN(0xfffc, 0, 0, 1)), 'C1:D5');
  // Five rows up from row 5 is the last row of the sheet; three columns left of C is the last but one column.
  assert.equal(text(ptg.refN(0xfffb, 0xfd)), 'IV65536');
});

test('functions: fixed and variable argument counts, SUM as an attribute, IF and CHOOSE with their jump marks', () => {
  assert.equal(text(tokens(ptg.func(19))), 'PI()');
  assert.equal(text(tokens(ptg.ref(0, 0), ptg.int(2), ptg.func(27))), 'ROUND(A1,2)');
  assert.equal(text(tokens(ptg.ref(0, 0), ptg.int(1), ptg.int(3), ptg.func(31, 0x21))), 'MID(A1,1,3)');
  assert.equal(text(tokens(ptg.area(0, 9, 0, 0), ptg.attrSum)), 'SUM(A1:A10)');
  assert.equal(text(tokens(ptg.area(0, 9, 0, 0), ptg.area(0, 9, 1, 1), ptg.funcVar(4, 2, 0x22))), 'SUM(A1:A10,B1:B10)');
  assert.equal(text(tokens(ptg.funcVar(0, 0))), 'COUNT()');
  assert.equal(text(tokens(ptg.attrVolatile, ptg.func(74))), 'NOW()');
  assert.equal(text(tokens(ptg.ref(0, 0), ptg.int(0), ptg.gt, ptg.attrIf(7), ptg.str('artı'), ptg.attrGoto(10), ptg.str('eksi'), ptg.attrGoto(3), ptg.funcVar(1, 3))), 'IF(A1>0,"artı","eksi")');
  assert.equal(text(tokens(ptg.ref(0, 0), ptg.attrChoose([6, 13, 20]), ptg.int(10), ptg.attrGoto(14), ptg.int(20), ptg.attrGoto(7), ptg.funcVar(100, 3))), 'CHOOSE(A1,10,20)');
  assert.equal(text(tokens(ptg.attrSpace(0, 2), ptg.ref(0, 0), ptg.attrSpace(1, 1), ptg.int(1), ptg.add)), 'A1+1');
  assert.equal(text(tokens(ptg.ref(1, 0), ptg.area(0, 9, 3, 5, false), ptg.int(2), ptg.bool(false), ptg.funcVar(102, 4))), 'VLOOKUP(A2,$D$1:$F$10,2,FALSE)');
  // The prompt bit of the argument count is not part of the count.
  assert.equal(text(tokens(ptg.int(1), [0x42, 0x81, 4, 0])), 'SUM(1)');
});

test('the function table: names by index, and fixed argument counts', () => {
  const expected: [number, string, number | undefined][] = [
    [0, 'COUNT', undefined],
    [1, 'IF', undefined],
    [4, 'SUM', undefined],
    [10, 'NA', 0],
    [15, 'SIN', 1],
    [27, 'ROUND', 2],
    [31, 'MID', 3],
    [65, 'DATE', 3],
    [74, 'NOW', 0],
    [82, 'SEARCH', undefined],
    [97, 'ATAN2', 2],
    [102, 'VLOOKUP', undefined],
    [109, 'LOG', undefined],
    [119, 'REPLACE', 4],
    [124, 'FIND', undefined],
    [148, 'INDIRECT', undefined],
    [169, 'COUNTA', undefined],
    [183, 'PRODUCT', undefined],
    [197, 'TRUNC', undefined],
    [212, 'ROUNDUP', 2],
    [221, 'TODAY', 0],
    [227, 'MEDIAN', undefined],
    [228, 'SUMPRODUCT', undefined],
    [247, 'DB', undefined],
    [252, 'FREQUENCY', 2],
    [261, 'ERROR.TYPE', 1],
    [269, 'AVEDEV', undefined],
    [288, 'CEILING', 2],
    [302, 'WEIBULL', 4],
    [325, 'LARGE', 2],
    [336, 'CONCATENATE', undefined],
    [337, 'POWER', 2],
    [344, 'SUBTOTAL', undefined],
    [345, 'SUMIF', undefined],
    [346, 'COUNTIF', 2],
    [347, 'COUNTBLANK', 1],
    [351, 'DATEDIF', 3],
    [359, 'HYPERLINK', undefined],
    [367, 'VARA', undefined],
    [379, 'RTD', undefined],
    [480, 'IFERROR', undefined],
    [484, 'AVERAGEIFS', undefined],
  ];
  for (const [index, name, count] of expected) {
    assert.equal(functionName(index), name);
    assert.equal(fixedArgCount(index), count, name);
  }
  for (const gap of [202, 203, 217, 218, 249, 250, 255, 333, 485, 1000]) assert.equal(functionName(gap), undefined);
  // No name appears twice.
  const seen = new Set<string>();
  for (let i = 0; i < 500; i++) {
    const name = functionName(i);
    if (name === undefined) continue;
    assert.ok(!seen.has(name), name);
    seen.add(name);
  }
  assert.equal(seen.size, 477);
});

test('names, names of other sheets and workbooks, add-in and later functions', () => {
  assert.equal(text(ptg.name(1)), 'Toplam');
  assert.equal(text(tokens(ptg.name(2), ptg.int(100), ptg.mul)), 'KDV_Oranı*100');
  // A function Excel 97 does not know is a name, called through the "user defined" entry.
  assert.equal(text(tokens(ptg.name(3), ptg.ref(0, 0), ptg.int(0), ptg.funcVar(255, 3))), '_xlfn.IFERROR(A1,0)');
  assert.equal(text(tokens(ptg.nameX(6, 1), ptg.ref(0, 0), ptg.int(1), ptg.funcVar(255, 3))), 'EOMONTH(A1,1)');
  assert.equal(text(ptg.nameX(7, 1)), 'Toplam');
  assert.equal(text(ptg.nameX(1, 1)), "'Veri sayfası'!Toplam");
  assert.equal(text(ptg.nameX(8, 1)), '[1]!Kur');
  assert.equal(text(ptg.nameX(4, 1)), '[1]Data!Kur');
  assert.equal(text(ptg.nameX(5, 1)), "'[1]Dış sayfa'!Kur");
});

test('3D references: sheet names, quoting, sheet ranges, other workbooks', () => {
  assert.equal(text(ptg.ref3d(0, 0, 0)), 'Sheet1!$A$1');
  assert.equal(text(ptg.ref3d(1, 4, 1, true)), "'Veri sayfası'!B5");
  assert.equal(text(ptg.ref3d(9, 0, 0)), "'Ali''nin'!$A$1");
  // A sheet name that looks like a cell reference has to be quoted.
  assert.equal(text(ptg.ref3d(10, 0, 0)), "'Q1'!$A$1");
  assert.equal(text(ptg.area3d(1, 0, 9, 0, 1)), "'Veri sayfası'!$A$1:$B$10");
  assert.equal(text(tokens(ptg.area3d(2, 0, 9, 0, 0), ptg.funcVar(4, 1))), "SUM('Sheet1:Q1'!$A$1:$A$10)");
  assert.equal(text(ptg.ref3d(4, 0, 0)), '[1]Data!$A$1');
  assert.equal(text(ptg.area3d(5, 0, 1, 0, 0)), "'[1]Dış sayfa'!$A$1:$A$2");
  assert.equal(text(ptg.area3d(0, 0, 65535, 0, 0)), 'Sheet1!$A:$A');
  assert.equal(text(ptg.refErr3d(0)), '#REF!');
  assert.equal(quoteSheetName('Sheet_1.b'), 'Sheet_1.b');
  assert.equal(quoteSheetName('2024'), "'2024'");
  assert.equal(quoteSheetName('RC'), "'RC'");
  assert.equal(quoteSheetName('TRUE'), "'TRUE'");
  assert.equal(quoteSheetName('Ürünler'), "'Ürünler'");
});

test('array constants are read from the data after the tokens', () => {
  // {1,"iki";TRUE,#N/A}, then {-2.5}.
  const first = [1, ...u16(1), 0x01, ...f64(1), 0x02, ...str16('iki'), 0x04, 1, 0, 0, 0, 0, 0, 0, 0, 0x10, 0x2a, 0, 0, 0, 0, 0, 0, 0];
  const second = [0, ...u16(0), 0x01, ...f64(-2.5)];
  assert.equal(text(tokens(ptg.array(), ptg.array(0x20), ptg.funcVar(4, 2))), undefined);
  assert.equal(text(tokens(ptg.array(), ptg.array(0x40), ptg.funcVar(4, 2)), [...first, ...second]), 'SUM({1,"iki";TRUE,#N/A},{-2.5})');
  // The rectangles of a PtgMemArea come before the array of the token after it.
  assert.equal(text(tokens(ptg.memArea(0), ptg.array()), [...u16(1), 0, 0, 0, 0, 0, 0, 0, 0, ...second]), '{-2.5}');
  // An empty element cannot be written.
  assert.equal(text(ptg.array(), [0, 0, 0, 0x00, 0, 0, 0, 0, 0, 0, 0, 0]), undefined);
});

test('PtgMemFunc, PtgMemArea, PtgMemErr: the reference they stand for follows as tokens', () => {
  const union = tokens(ptg.area3d(0, 0, 65535, 0, 1), ptg.area3d(0, 0, 1, 0, 255), ptg.union);
  assert.equal(text(tokens(ptg.memFunc(union.length), union)), 'Sheet1!$A:$B,Sheet1!$1:$2');
  assert.equal(text(tokens(ptg.memErr(0), ptg.ref(0, 0))), 'A1');
});

test('tokens that cannot be turned into text give no text, with the reason counted', () => {
  const ctx = context();
  const failing: [number[], RegExp][] = [
    [[0x18, 1, 2, 3, 4, 5], /0x18/],
    [[0x7f], /0x7f/],
    [[0x80], /0x80/],
    [tokens(ptg.int(1), ptg.int(2)), /left over/],
    [ptg.add, /without operands/],
    [[], /empty/],
    [ptg.missArg, /empty/],
    [ptg.func(4), /no fixed argument count/],
    [ptg.func(999), /no fixed argument count/],
    [tokens(ptg.int(1), ptg.funcVar(999, 1)), /unknown function 999/],
    [tokens(ptg.int(1), ptg.funcVar(4 | 0x8000, 1)), /command equivalent/],
    [ptg.funcVar(4, 3), /too few operands/],
    [ptg.name(9), /unknown name/],
    [ptg.nameX(99, 1), /no EXTERNSHEET/],
    [ptg.nameX(8, 2), /unknown external name/],
    [ptg.ref3d(3, 0, 0), /deleted or unknown sheet/],
    [ptg.ref3d(6, 0, 0), /add-in/],
    [[0x1e, 1], /truncated/],
    [[0x17, 200, 0, 65], /truncated/],
    [ptg.err(0x55), /unknown error/],
    [ptg.num(NaN), /not a number/],
    [tokens(ptg.str('1+1'), ptg.int(1), ptg.funcVar(255, 2)), /unusable name/],
    [tokens(ptg.exp(0, 0)), /PtgExp/],
    [[0x19, 0x04, 0xff, 0xff], /truncated/],
  ];
  for (const [parts, reason] of failing) {
    const before = new Map(ctx.failures);
    assert.equal(text(parts, [], ctx), undefined);
    const added = [...ctx.failures].filter(([key, value]) => before.get(key) !== value).map(([key]) => key);
    assert.equal(added.length, 1);
    assert.match(added[0], reason);
  }
});

test('BIFF5 tokens: narrower references, byte strings, sheet indexes inside the token', () => {
  const ctx = context({ biff8: false, codepage: 1254, supBooks: [{ kind: 'self', sheetNames: ['Veri sayfası'], externNames: [], ordinal: 0 }, { kind: 'addin', sheetNames: [], externNames: ['XIRR'], ordinal: 0 }] });
  const t5 = (parts: number[], extra: number[] = []) => text(parts, extra, ctx);
  assert.equal(t5([0x44, ...u16(9 | 0xc000), 3]), 'D10');
  assert.equal(t5([0x44, ...u16(9), 3]), '$D$10');
  assert.equal(t5([0x25, ...u16(0), ...u16(16383), 1, 2]), '$B:$C');
  assert.equal(t5([0x25, ...u16(0 | 0xc000), ...u16(4 | 0xc000), 0, 1]), 'A1:B5');
  assert.equal(t5([0x17, 2, 0xfe, 0xf0]), '"şğ"');
  // Relative: one row up (14 bits), one column left, from C5.
  assert.equal(t5([0x4c, ...u16(0x3fff | 0xc000), 0xff]), 'B4');
  assert.equal(t5([0x43, ...u16(1), ...new Array(12).fill(0)]), 'Toplam');
  const sheetRef = (first: number, last: number) => [...u16(0xffff), 0, 0, 0, 0, 0, 0, 0, 0, ...u16(first), ...u16(last)];
  assert.equal(t5([0x3a, ...sheetRef(1, 1), ...u16(0), 0]), "'Veri sayfası'!$A$1");
  assert.equal(t5([0x3b, ...sheetRef(0, 3), ...u16(0), ...u16(1), 0, 0]), "'Sheet1:Q1'!$A$1:$A$2");
  assert.equal(t5([0x3a, ...u16(1), 0, 0, 0, 0, 0, 0, 0, 0, ...u16(0), ...u16(0), ...u16(0), 0]), "'Veri sayfası'!$A$1");
  assert.equal(t5([0x3a, ...u16(2), 0, 0, 0, 0, 0, 0, 0, 0, ...u16(0), ...u16(0), ...u16(0), 0]), undefined);
  assert.equal(t5([0x39, ...u16(2), 0, 0, 0, 0, 0, 0, 0, 0, ...u16(1), ...new Array(12).fill(0), 0x1e, 1, 0, 0x42, 2, 255, 0]), 'XIRR(1)');
  assert.equal(t5([0x60, 0, 0, 0, 0, 0, 0, 0], [2, ...u16(1), 0x01, ...f64(1), 0x02, 1, 0xfd]), '{1,"ı"}');
});

const twoSheets = (records: Uint8Array, tail?: Uint8Array) =>
  readXlsStream(workbookStream({ tail, sheets: [{ name: 'Sheet1', records }, { name: 'Veri sayfası', records: cat(number(0, 0, 10), number(1, 0, 20)) }] }));

test('shared formulas: every cell of the group gets the text for its own position', () => {
  // C1:C3 = A1+$B$1, A2+$B$1, A3+$B$1.
  const sharedTokens = tokens(ptg.refN(0, 0xfe), ptg.ref(0, 1, false, false), ptg.add);
  const records = cat(
    number(0, 0, 1),
    number(1, 0, 2),
    number(2, 0, 3),
    number(0, 1, 100),
    formula(0, 2, ptg.exp(0, 2), 101),
    shrFmla(0, 2, 2, 2, sharedTokens),
    formula(1, 2, ptg.exp(0, 2), 102),
    formula(2, 2, ptg.exp(0, 2), { text: 'yüz üç' }),
    // A cell that points to a group that does not exist keeps its value and has no formula.
    formula(5, 5, ptg.exp(9, 9), 7),
  );
  const { workbook, report } = twoSheets(records);
  const sheet = workbook.sheets[0];
  assert.deepEqual(getCell(sheet, 0, 2)?.formula, { text: 'A1+$B$1', type: 'shared', si: 0, ref: 'C1:C3' });
  // The cells after the first point to it, as in an xlsx file; written out, each has the text for its own position.
  assert.deepEqual(getCell(sheet, 1, 2)?.formula, { text: '', type: 'shared', si: 0 });
  assert.deepEqual(getCell(sheet, 2, 2)?.formula, { text: '', type: 'shared', si: 0 });
  const masters = sharedMasters(sheet);
  assert.deepEqual([1, 2].map((row) => explicitFormula(getCell(sheet, row, 2)!, row, 2, masters).formula), [
    { text: 'A2+$B$1', type: 'normal' },
    { text: 'A3+$B$1', type: 'normal' },
  ]);
  assert.equal(getCell(sheet, 2, 2)?.value, 'yüz üç');
  assert.deepEqual(getCell(sheet, 5, 5), { kind: 'number', value: 7, style: 0 });
  assert.equal(report.untranslatedFormulas, 1);
  assert.equal(report.untranslatedReasons.get('PtgExp without its group'), 1);
  // The engine calculates the same values the file holds for the first two.
  const engine = new CalcEngine(workbook);
  engine.recalculate(undefined, { overwrite: true });
  assert.deepEqual([0, 1, 2].map((row) => getCell(workbook.sheets[0], row, 2)?.value), [101, 102, 103]);
});

test('array formulas and data tables: the first cell of the range carries the formula', () => {
  const records = cat(
    formula(0, 3, ptg.exp(0, 3), 3),
    arrayRec(0, 1, 3, 3, tokens(ptg.area(0, 1, 0, 0), ptg.array(), ptg.mul), [1, ...u16(0), 0x01, ...f64(2), 0x01, ...f64(3)]),
    formula(1, 3, ptg.exp(0, 3), 6),
    formula(4, 1, ptg.tbl(4, 1), 5),
    rec(0x0236, u16(4), u16(6), [1, 2], u16(0x0008 | 0x0004), u16(0), u16(0), u16(1), u16(0)),
    formula(4, 2, ptg.tbl(4, 1), 6),
  );
  const { workbook, report } = twoSheets(records);
  const sheet = workbook.sheets[0];
  assert.deepEqual(getCell(sheet, 0, 3)?.formula, { text: 'A1:A2*{2,3}', type: 'array', ref: 'D1:D2' });
  assert.deepEqual(getCell(sheet, 1, 3), { kind: 'number', value: 6, style: 0 });
  const table = getCell(sheet, 4, 1)?.formula;
  assert.equal(table?.type, 'dataTable');
  assert.equal(table?.ref, 'B5:C7');
  assert.deepEqual([...(table?.attrs ?? [])], [['dt2D', '1'], ['dtr', '1'], ['r1', 'A1'], ['r2', 'A2']]);
  assert.equal(getCell(sheet, 4, 2)?.formula, undefined);
  assert.equal(report.untranslatedFormulas, 0);
});

test('a formula that cannot be translated keeps its value and gets no formula', () => {
  const { workbook, report } = twoSheets(cat(formula(0, 0, [0x18, 1, 2, 3, 4, 5], 42), formula(0, 1, tokens(ptg.int(1), ptg.int(1), ptg.add), 2)));
  assert.deepEqual(getCell(workbook.sheets[0], 0, 0), { kind: 'number', value: 42, style: 0 });
  assert.equal(getCell(workbook.sheets[0], 0, 1)?.formula?.text, '1+1');
  assert.equal(report.untranslatedFormulas, 1);
});

test('3D references and defined names of a workbook, through EXTERNSHEET, SUPBOOK and NAME', () => {
  const tail = cat(
    supBookSelf(2),
    supBookExternal('\u0001C:\u0003veri\u0003kur.xls', ['Kur']),
    externName('Günlük'),
    supBookAddin(),
    externName('EOMONTH'),
    externSheet([0, 1, 1], [0, 0, 0], [1, 0, 0], [2, -2, -2], [0, -2, -2], [1, -2, -2]),
    definedName({ name: 'Toplam', tokens: ptg.area3d(0, 0, 1, 0, 0) }),
    definedName({ name: '', builtin: 6, sheet: 1, tokens: ptg.area3d(1, 0, 3, 0, 2) }),
    definedName({ name: '', builtin: 7, sheet: 2, hidden: true, tokens: tokens(ptg.memFunc(23), ptg.area3d(0, 0, 65535, 0, 0), ptg.area3d(0, 0, 0, 0, 255), ptg.union) }),
    definedName({ name: '', builtin: 13, sheet: 1, hidden: true, tokens: ptg.area3d(1, 0, 9, 0, 3) }),
    // Refers to a name defined after it.
    definedName({ name: 'Oran', tokens: tokens(ptg.name(6), ptg.int(2), ptg.mul) }),
    definedName({ name: 'Taban', tokens: ptg.num(0.09) }),
    // No formula: an add-in function or a macro; it is not a defined name of the model.
    definedName({ name: '_xlfn.IFERROR', tokens: [] }),
    definedName({ name: 'Bozuk', tokens: [0x18, 1, 2, 3, 4, 5] }),
  );
  const records = cat(
    formula(0, 0, tokens(ptg.area3d(0, 0, 1, 0, 0), ptg.funcVar(4, 1)), 30),
    formula(1, 0, tokens(ptg.name(1), ptg.funcVar(4, 1), ptg.name(5), ptg.mul), 5.4),
    formula(2, 0, tokens(ptg.ref3d(2, 0, 0), ptg.nameX(5, 1), ptg.mul), 1),
    formula(3, 0, tokens(ptg.nameX(3, 1), ptg.ref(0, 0), ptg.int(1), ptg.funcVar(255, 3)), 2),
    formula(4, 0, tokens(ptg.name(7), ptg.ref(0, 0), ptg.int(0), ptg.funcVar(255, 3)), 30),
    formula(5, 0, tokens(ptg.nameX(4, 1), ptg.funcVar(4, 1)), 30),
  );
  const { workbook, report } = twoSheets(records, tail);
  const sheet = workbook.sheets[0];
  assert.deepEqual([0, 1, 2, 3, 4, 5].map((row) => getCell(sheet, row, 0)?.formula?.text), [
    "SUM('Veri sayfası'!$A$1:$A$2)",
    'SUM(Toplam)*Oran',
    '[1]Kur!$A$1*[1]!Günlük',
    'EOMONTH(A1,1)',
    '_xlfn.IFERROR(A1,0)',
    'SUM(Toplam)',
  ]);
  assert.deepEqual(workbook.definedNames, [
    { name: 'Toplam', value: "'Veri sayfası'!$A$1:$A$2", localSheet: undefined, hidden: false },
    { name: '_xlnm.Print_Area', value: 'Sheet1!$A$1:$C$4', localSheet: 0, hidden: false },
    { name: '_xlnm.Print_Titles', value: "'Veri sayfası'!$A:$A,'Veri sayfası'!$1:$1", localSheet: 1, hidden: true },
    { name: '_xlnm._FilterDatabase', value: 'Sheet1!$A$1:$D$10', localSheet: 0, hidden: true },
    { name: 'Oran', value: 'Taban*2', localSheet: undefined, hidden: false },
    { name: 'Taban', value: '0.09', localSheet: undefined, hidden: false },
  ]);
  assert.equal(report.untranslatedNames, 1);
  assert.equal(report.untranslatedFormulas, 0);

  // What the reader writes, the formula parser of the engine reads, and the engine calculates.
  for (const row of [0, 1, 2, 3, 4, 5]) assert.doesNotThrow(() => parseFormula(getCell(sheet, row, 0)?.formula?.text ?? ''));
  // A union outside parentheses (print titles) is written as xlsx files write it; the engine's parser does not take that form.
  for (const name of workbook.definedNames) if (name.name !== '_xlnm.Print_Titles') assert.doesNotThrow(() => parseFormula(name.value));
  const engine = new CalcEngine(workbook);
  engine.recalculate(undefined, { overwrite: true });
  assert.equal(getCell(workbook.sheets[0], 0, 0)?.value, 30);
  assert.ok(Math.abs((getCell(workbook.sheets[0], 1, 0)?.value as number) - 5.4) < 1e-12);
  assert.equal(getCell(workbook.sheets[0], 5, 0)?.value, 30);
});

test('a formula whose tokens go on in a CONTINUE record', () => {
  const parts = tokens(ptg.str('a'.repeat(200)), ptg.str('b'.repeat(200)), ptg.concat);
  const whole = formula(0, 0, parts, { text: '' });
  // Cut the FORMULA record in two: the second half becomes a CONTINUE record.
  const header = 4 + 22;
  const first = rec(0x0006, whole.subarray(4, header + 150));
  const second = rec(0x003c, whole.subarray(header + 150, header + parts.length));
  const { workbook } = twoSheets(cat(first, second, rec(0x0207, str16('sonuç'))));
  const cell = getCell(workbook.sheets[0], 0, 0);
  assert.equal(cell?.formula?.text, `"${'a'.repeat(200)}"&"${'b'.repeat(200)}"`);
  assert.equal(cell?.value, 'sonuç');
  assert.ok(narrow);
});
