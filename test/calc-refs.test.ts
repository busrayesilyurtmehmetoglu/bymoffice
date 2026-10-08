import assert from 'node:assert/strict';
import { test } from 'node:test';
import { adjustAnchoredFormula, adjustFormula, adjustRange, adjustReferenceText, relativeExtent, shiftFormula, type StructuralEdit } from '../src/calc/refs';
import { MAX_COLS, MAX_ROWS } from '../src/model/sheet/types';

const SHEET = 'Sheet1';

const edit = (kind: 'insert' | 'delete', axis: 'row' | 'col', at: number, count: number, sheetName = SHEET): StructuralEdit => ({
  kind,
  axis,
  at,
  count,
  sheetName,
});

/** Rows of [input, dRow, dCol, expected]. */
function checkShift(cases: [string, number, number, string][]): void {
  for (const [input, dRow, dCol, expected] of cases) {
    assert.equal(shiftFormula(input, dRow, dCol), expected, `shift ${JSON.stringify(input)} by (${dRow}, ${dCol})`);
  }
}

/** Rows of [input, expected], all for the same edit and a formula on Sheet1. */
function checkAdjust(e: StructuralEdit, cases: [string, string][], formulaSheet = SHEET): void {
  for (const [input, expected] of cases) {
    assert.equal(adjustFormula(input, formulaSheet, e), expected, `${e.kind} ${e.axis} ${e.at}+${e.count}: ${JSON.stringify(input)}`);
  }
}

// ---------------------------------------------------------------------------
// shiftFormula

test('refs shift: cells, anchors and letter case', () => {
  checkShift([
    ['A1', 1, 1, 'B2'],
    ['$A$1', 1, 1, '$A$1'],
    ['A$1', 1, 1, 'B$1'],
    ['$A1', 1, 1, '$A2'],
    ['B2', -1, -1, 'A1'],
    ['Z9', 1, 1, 'AA10'],
    ['XFC1048575', 1, 1, 'XFD1048576'],
    // A changed reference is written in upper case; an unchanged one keeps its spelling.
    ['a1', 1, 0, 'A2'],
    ['$a$1+b2', 1, 0, '$a$1+B3'],
    ['a$1', 5, 0, 'a$1'],
    ['=A1', 1, 0, '=A2'],
    ['-A1^2+A1%', 1, 0, '-A2^2+A2%'],
    ['@A1', 0, 1, '@B1'],
    ['A1#', 1, 0, 'A2#'],
  ]);
});

test('refs shift: ranges, whole columns and rows, operators', () => {
  checkShift([
    ['A1:B2', 1, 1, 'B2:C3'],
    ['$A$1:B2', 1, 1, '$A$1:C3'],
    ['A1:$B$2', 2, 0, 'A3:$B$2'],
    ['A:A', 5, 1, 'B:B'],
    ['A:A', 5, 0, 'A:A'],
    ['$A:$C', 3, 3, '$A:$C'],
    ['$A:C', 0, 1, '$A:D'],
    ['1:1', 2, 5, '3:3'],
    ['1:1', 0, 5, '1:1'],
    ['$1:$3', 5, 0, '$1:$3'],
    ['1:$3', 1, 0, '2:$3'],
    ['SUM(A:A)+SUM(2:2)', 1, 1, 'SUM(B:B)+SUM(3:3)'],
    // Intersection (space) and union (comma) operands move one by one.
    ['A1:B5 B2:C3', 1, 0, 'A2:B6 B3:C4'],
    ['SUM((A1,B2,$C$3))', 1, 1, 'SUM((B2,C3,$C$3))'],
    ['A1:INDEX(B:B,5)', 1, 0, 'A2:INDEX(B:B,5)'],
    ['IF(A1>0,SUM(B1:B10)/COUNT($C$1:C10),"")', 2, 1, 'IF(B3>0,SUM(C3:C12)/COUNT($C$1:D12),"")'],
    ['ROUND(IF(ISERROR(VLOOKUP(A2,$D$2:$E$9,2,0)),0,VLOOKUP(A2,$D$2:$E$9,2,0)),2)', 3, 0, 'ROUND(IF(ISERROR(VLOOKUP(A5,$D$2:$E$9,2,0)),0,VLOOKUP(A5,$D$2:$E$9,2,0)),2)'],
  ]);
});

test('refs shift: off the sheet becomes #REF!', () => {
  checkShift([
    ['A1', -1, 0, '#REF!'],
    ['A1', 0, -1, '#REF!'],
    ['XFD1', 0, 1, '#REF!'],
    ['A1048576', 1, 0, '#REF!'],
    ['$A1', -1, 0, '#REF!'],
    ['A$1', -1, 0, 'A$1'],
    // One corner off the sheet takes the whole range with it.
    ['A1:B2', 0, -1, '#REF!'],
    ['A1:B1048576', 1, 0, '#REF!'],
    ['SUM(A1:B2,C3)', -1, 0, 'SUM(#REF!,C2)'],
    ['Sheet1!A1', -1, 0, 'Sheet1!#REF!'],
    ["'My Sheet'!A1:B2", -1, 0, "'My Sheet'!#REF!"],
    ['Sheet1:Sheet3!A1', 0, -1, 'Sheet1:Sheet3!#REF!'],
    ['A:B', 0, -1, '#REF!'],
    ['XFD:XFD', 0, 1, '#REF!'],
    ['1:2', -1, 0, '#REF!'],
    ['A1', -MAX_ROWS, 0, '#REF!'],
    ['A1', 1e300, 0, '#REF!'],
  ]);
});

test('refs shift: sheet prefixes', () => {
  checkShift([
    ['Sheet1!A1', 1, 0, 'Sheet1!A2'],
    ['Sheet1!$A$1:B2', 1, 1, 'Sheet1!$A$1:C3'],
    ["'My Sheet'!A1", 1, 0, "'My Sheet'!A2"],
    ["'It''s'!A1", 1, 0, "'It''s'!A2"],
    ["'Veri Sayfası'!A1", 1, 1, "'Veri Sayfası'!B2"],
    ['Özet!B2', 1, 0, 'Özet!B3'],
    ['Özet!B2+Çıktı!$C3*Şube_İl!D$4', 1, 1, 'Özet!C3+Çıktı!$C4*Şube_İl!E$4'],
    ['Sheet1:Sheet3!A1', 1, 0, 'Sheet1:Sheet3!A2'],
    ["'Sheet 1:Sheet 3'!A1:B2", 1, 0, "'Sheet 1:Sheet 3'!A2:B3"],
    ['Sheet1!A:A', 0, 1, 'Sheet1!B:B'],
    ['Sheet1!1:1', 1, 0, 'Sheet1!2:2'],
    // Both corners prefixed with the same sheet: one range, prefixes kept.
    ['Sheet1!A1:Sheet1!B2', 1, 0, 'Sheet1!A2:Sheet1!B3'],
    ['Sheet1!A1:Sheet1!$B$2', -1, 0, 'Sheet1!#REF!'],
    ['Sheet1!A1:Sheet2!B2', 1, 0, 'Sheet1!A2:Sheet2!B3'],
    // References into other workbooks move like the others (a shared formula relies on it).
    ['[1]Sheet1!A1', 1, 0, '[1]Sheet1!A2'],
    ['[1]Sheet1!A1', 1, 1, '[1]Sheet1!B2'],
    ['[1]Sheet1!$A$1', 1, 1, '[1]Sheet1!$A$1'],
    ['[1]Sheet1!A1:B2+A1', 1, 1, '[1]Sheet1!B2:C3+B2'],
    ['[1]Sheet1:Sheet3!A1+A1', 1, 0, '[1]Sheet1:Sheet3!A2+A2'],
    ["'[Book.xlsx]Sheet1'!A1+A1", 1, 0, "'[Book.xlsx]Sheet1'!A2+A2"],
    ["'C:\\x\\[b.xlsx]S'!$B5", 1, 1, "'C:\\x\\[b.xlsx]S'!$B6"],
    ["'C:\\dir\\[Book.xlsx]Sheet1'!$A1+A1", 1, 0, "'C:\\dir\\[Book.xlsx]Sheet1'!$A2+A2"],
    ['[1]Sheet1!A1', -1, 0, '[1]Sheet1!#REF!'],
    ['[1]Sheet1!A:A+[1]Sheet1!2:2', 1, 1, '[1]Sheet1!B:B+[1]Sheet1!3:3'],
    ['[2]!Rates+A1', 1, 0, '[2]!Rates+A2'],
    // A sheet-scoped name is not a reference.
    ['Sheet1!Total+A1', 1, 0, 'Sheet1!Total+A2'],
  ]);
});

test('refs shift: text that only looks like a reference', () => {
  checkShift([
    // String constants.
    ['"A1"&A1', 1, 0, '"A1"&A2'],
    ['"say ""A1"" "&B1', 1, 0, '"say ""A1"" "&B2'],
    ['""&A1&""', 1, 0, '""&A2&""'],
    ["\"it's\"&A1&\"'x'!B1\"", 1, 0, "\"it's\"&A2&\"'x'!B1\""],
    // Function names.
    ['LOG10(A1)+ATAN2(B1,C1)', 1, 0, 'LOG10(A2)+ATAN2(B2,C2)'],
    ['DAYS360(A1,B1)', 1, 0, 'DAYS360(A2,B2)'],
    ['A1B(A1)', 1, 0, 'A1B(A2)'],
    ['log10(a1)', 1, 0, 'log10(A2)'],
    ['_xlfn.XLOOKUP(A1,B:B,C:C)', 1, 0, '_xlfn.XLOOKUP(A2,B:B,C:C)'],
    // Names. TAX2024 is column TAX, row 2024: a cell, as in Excel.
    ['Toplam+A1B2C+_A1+TAX2024', 1, 0, 'Toplam+A1B2C+_A1+TAX2025'],
    ['ABCD1+A1048577+XFE1+A0+AA', 1, 1, 'ABCD1+A1048577+XFE1+A0+AA'],
    ['XFD1048576', 0, 0, 'XFD1048576'],
    ['Satış.A1+A1.5+A1_x+x_A1+A1\\B', 1, 1, 'Satış.A1+A1.5+A1_x+x_A1+A1\\B'],
    // Three letters and a row are a cell whatever they spell.
    ['xA1+MAX9', 1, 1, 'XB2+MAY10'],
    ['ÇA1+A1Ç+İ5', 1, 1, 'ÇA1+A1Ç+İ5'],
    ['TRUE+FALSE+PI()+A', 1, 1, 'TRUE+FALSE+PI()+A'],
    // Error constants.
    ['#REF!+#N/A+#DIV/0!+#NAME?+#NUM!+#VALUE!+#NULL!+A1', 1, 0, '#REF!+#N/A+#DIV/0!+#NAME?+#NUM!+#VALUE!+#NULL!+A2'],
    ['Sheet1!#REF!+A1', 1, 0, 'Sheet1!#REF!+A2'],
    // What a reference to a deleted sheet turns into.
    ['#REF!A1+#REF!$B$2:C3+A1', 1, 0, '#REF!A1+#REF!$B$2:C3+A2'],
    // Structured references.
    ['Tablo1[Sütun]+[@Sütun]+A1', 1, 0, 'Tablo1[Sütun]+[@Sütun]+A2'],
    ['Tbl1[A1]+SUM(Tbl1[[#This Row],[B2]:[C3]])', 1, 1, 'Tbl1[A1]+SUM(Tbl1[[#This Row],[B2]:[C3]])'],
    ["T[It's '] A1]+A1", 1, 0, "T[It's '] A1]+A2"],
    // Numbers.
    ['1E5+2.5E-3+1.5e+10+E5', 1, 0, '1E5+2.5E-3+1.5e+10+E6'],
    ['A1*100+.5+3', 1, 0, 'A2*100+.5+3'],
    ['{1,2;3,4}*A1', 1, 0, '{1,2;3,4}*A2'],
    // R1C1 notation is not supported and stays.
    ['R1C1+R[1]C[-1]+RC[2]+R2C3:R4C5', 1, 1, 'R1C1+R[1]C[-1]+RC[2]+R2C3:R4C5'],
    // A mixed pair is not a range: only the cell moves.
    ['A1:B', 1, 0, 'A2:B'],
    ['A:B1', 1, 1, 'A:C2'],
  ]);
});

test('refs shift: everything around a reference is kept byte for byte', () => {
  const formula = '  sum( a1 ;  $B$2 ) &\t"x"\n+ If(c3,1 , 2)  ';
  assert.equal(shiftFormula(formula, 0, 0), formula);
  assert.equal(shiftFormula(formula, 1, 0), '  sum( A2 ;  $B$2 ) &\t"x"\n+ If(C4,1 , 2)  ');
  assert.equal(shiftFormula('', 3, 3), '');
  // Bad deltas move nothing.
  assert.equal(shiftFormula('A1', NaN, NaN), 'A1');
  assert.equal(shiftFormula('A1', 1.9, 0), 'A2');
});

// ---------------------------------------------------------------------------
// adjustFormula

test('refs adjust: inserting rows', () => {
  // Two rows inserted before row 3.
  checkAdjust(edit('insert', 'row', 2, 2), [
    ['A1', 'A1'],
    ['A2', 'A2'],
    ['A3', 'A5'],
    ['A4', 'A6'],
    // "$" does not hold a reference in place against a structural edit.
    ['$A$3', '$A$5'],
    ['A$3+$A3', 'A$5+$A5'],
    // Just after the end: no growth.
    ['A1:A2', 'A1:A2'],
    // Inside (start < at <= end): grows.
    ['A1:A3', 'A1:A5'],
    ['A2:A4', 'A2:A6'],
    ['$A$2:$C$4', '$A$2:$C$6'],
    // Exactly at the start: moves as a whole.
    ['A3:A5', 'A5:A7'],
    ['A10:B20', 'A12:B22'],
    // Corners written in reverse order stay reversed.
    ['A4:A2', 'A6:A2'],
    ['A:C', 'A:C'],
    ['3:3', '5:5'],
    ['1:3', '1:5'],
    ['$1:$2', '$1:$2'],
    ['a3', 'A5'],
    ['a2', 'a2'],
    ['SUM(A1:A3)+B3*"A3"', 'SUM(A1:A5)+B5*"A3"'],
    ['=SUM(A1:A3)', '=SUM(A1:A5)'],
    ['A1:A3 A3:C3', 'A1:A5 A5:C5'],
    ['TAX2024+LOG10(A3)', 'TAX2026+LOG10(A5)'],
  ]);
});

test('refs adjust: inserting columns', () => {
  checkAdjust(edit('insert', 'col', 1, 1), [
    ['A1', 'A1'],
    ['B1', 'C1'],
    ['$B$1', '$C$1'],
    ['A1:C1', 'A1:D1'],
    ['A1:A9', 'A1:A9'],
    ['B1:C9', 'C1:D9'],
    ['B:B', 'C:C'],
    ['A:C', 'A:D'],
    ['A:A', 'A:A'],
    ['1:1', '1:1'],
    ['Z1', 'AA1'],
  ]);
});

test('refs adjust: insertion at the end of the sheet', () => {
  const last = MAX_ROWS;
  checkAdjust(edit('insert', 'row', 0, 3), [
    [`A${last}`, '#REF!'],
    [`A${last - 2}`, '#REF!'],
    [`A${last - 3}`, `A${last}`],
    // The part pushed past the last row is cut off.
    [`A${last - 6}:A${last}`, `A${last - 3}:A${last}`],
    [`A${last - 1}:A${last}`, '#REF!'],
    [`Sheet1!A${last}`, 'Sheet1!#REF!'],
    // A reference to every row is "all rows" and stays that (see the whole-axis tests below).
    [`1:${last}`, `1:${last}`],
  ]);
  checkAdjust(edit('insert', 'row', 5, 3), [[`A1:A${last}`, `A1:A${last}`]]);
  checkAdjust(edit('insert', 'col', 0, 1), [
    ['XFD1', '#REF!'],
    ['XFC1', 'XFD1'],
    ['XFC1:XFD2', 'XFD1:XFD2'],
    ['XFD:XFD', '#REF!'],
  ]);
});

test('refs adjust: deleting rows', () => {
  // Rows 3 and 4 deleted.
  checkAdjust(edit('delete', 'row', 2, 2), [
    ['A1', 'A1'],
    ['A2', 'A2'],
    ['A3', '#REF!'],
    ['A4', '#REF!'],
    ['$A$4', '#REF!'],
    ['A5', 'A3'],
    ['$A$5', '$A$3'],
    ['A1:A2', 'A1:A2'],
    // Tail deleted.
    ['A1:A3', 'A1:A2'],
    ['A1:A4', 'A1:A2'],
    // Middle deleted.
    ['A1:A10', 'A1:A8'],
    // Head deleted.
    ['A3:A6', 'A3:A4'],
    ['A4:A10', 'A3:A8'],
    // All of it deleted.
    ['A3:A4', '#REF!'],
    ['A3:C3', '#REF!'],
    ['A5:A6', 'A3:A4'],
    // UNSURE: a range shrunk to one cell is still written as a range ("A2:A2").
    ['A2:A3', 'A2:A2'],
    ['3:4', '#REF!'],
    ['2:5', '2:3'],
    ['A:A', 'A:A'],
    ['SUM(A3,B5)', 'SUM(#REF!,B3)'],
    ['Sheet1!A3', 'Sheet1!#REF!'],
    ["'Sheet1'!A3:B4+1", "'Sheet1'!#REF!+1"],
    ['A6:A2', 'A4:A2'],
  ]);
});

test('refs adjust: deleting columns', () => {
  checkAdjust(edit('delete', 'col', 1, 2), [
    ['A1', 'A1'],
    ['B1', '#REF!'],
    ['C5', '#REF!'],
    ['D1', 'B1'],
    ['A1:D1', 'A1:B1'],
    ['B1:C9', '#REF!'],
    ['C1:E1', 'B1:C1'],
    ['B:C', '#REF!'],
    ['A:D', 'A:B'],
    ['E:E', 'C:C'],
    ['1:5', '1:5'],
  ]);
});

test('refs adjust: only references to the edited sheet change', () => {
  const e = edit('delete', 'row', 2, 2);
  checkAdjust(e, [
    ['Sheet2!A5', 'Sheet2!A5'],
    ['Sheet1!A5', 'Sheet1!A3'],
    ['sheet1!A5', 'sheet1!A3'],
    ['SHEET1!$A$5', 'SHEET1!$A$3'],
    ["'Sheet1'!A5", "'Sheet1'!A3"],
    ['A5+Sheet2!A5', 'A3+Sheet2!A5'],
    ['[1]Sheet1!A5', '[1]Sheet1!A5'],
    ["'[Book.xlsx]Sheet1'!A5", "'[Book.xlsx]Sheet1'!A5"],
    // Structural edits never reach into another workbook, whatever its sheets are called.
    ['[1]Sheet1!A3:A9+[1]Sheet1!$A$3', '[1]Sheet1!A3:A9+[1]Sheet1!$A$3'],
    ["'C:\\x\\[b.xlsx]Sheet1'!A3:Sheet1!A9", "'C:\\x\\[b.xlsx]Sheet1'!A3:Sheet1!A7"],
    // 3D: adjusted when the edited sheet is one of the two ends.
    ['Sheet1:Sheet3!A5', 'Sheet1:Sheet3!A3'],
    ['Sheet0:Sheet1!A5', 'Sheet0:Sheet1!A3'],
    ["'Sheet0:Sheet1'!A5", "'Sheet0:Sheet1'!A3"],
    ['Sheet2:Sheet3!A5', 'Sheet2:Sheet3!A5'],
  ]);
  // A formula on another sheet: bare references are not about the edited sheet.
  checkAdjust(
    e,
    [
      ['A5', 'A5'],
      ['A5+Sheet1!A5', 'A5+Sheet1!A3'],
      ['SUM(A3:A4)+SUM(Sheet1!A3:A4)', 'SUM(A3:A4)+SUM(Sheet1!#REF!)'],
    ],
    'Sheet2',
  );
  // 3D reference whose ends do not name the edited sheet: left alone, even if it lies between them.
  checkAdjust(edit('delete', 'row', 2, 2, 'Sheet2'), [['Sheet1:Sheet3!A5', 'Sheet1:Sheet3!A5']], 'Sheet4');
});

test('refs adjust: Turkish and quoted sheet names', () => {
  const e = edit('insert', 'row', 0, 1, 'Veri Sayfası');
  checkAdjust(
    e,
    [
      ["'Veri Sayfası'!A1", "'Veri Sayfası'!A2"],
      ["'VERİ SAYFASI'!A1", "'VERİ SAYFASI'!A2"],
      ["'VERI SAYFASI'!A1", "'VERI SAYFASI'!A2"],
      ["'veri sayfası'!A1", "'veri sayfası'!A2"],
      ['Özet!B2', 'Özet!B2'],
      ['A1', 'A1'],
    ],
    'Özet',
  );
  checkAdjust(edit('insert', 'row', 0, 1, 'ÖZET'), [['özet!B2+Özet!B2+B2', 'özet!B3+Özet!B3+B3']], 'Özet');
  checkAdjust(edit('insert', 'col', 0, 1, "It's"), [["'It''s'!A1+'Its'!A1", "'It''s'!B1+'Its'!A1"]], 'Other');
});

test('refs adjust: edits that move nothing', () => {
  const formula = 'SUM(a1:$B$9)+Sheet2!C3';
  for (const e of [edit('insert', 'row', 0, 0), edit('delete', 'row', 0, -1), edit('insert', 'col', -1, 1), edit('delete', 'col', NaN, 1), edit('insert', 'row', 0, NaN)]) {
    assert.equal(adjustFormula(formula, SHEET, e), formula);
    assert.equal(adjustReferenceText('A1 B2', SHEET, e), 'A1 B2');
    assert.deepEqual(adjustRange({ top: 1, left: 1, bottom: 2, right: 2 }, e), { top: 1, left: 1, bottom: 2, right: 2 });
  }
  // References the edit does not reach keep their spelling.
  assert.equal(adjustFormula(formula, SHEET, edit('insert', 'row', 100, 5)), formula);
  assert.equal(adjustFormula(formula, SHEET, edit('delete', 'col', 100, 5)), formula);
});

test('refs adjust: sheet names compare with I, İ, ı and i folded together', () => {
  const del = (sheetName: string) => edit('delete', 'row', 0, 2, sheetName);
  checkAdjust(del('İşlem'), [["'işlem'!A3+İşlem!A3+İŞLEM!A3+ışlem!A3+Islem!A3", "'işlem'!A1+İşlem!A1+İŞLEM!A1+ışlem!A1+Islem!A3"]], 'Other');
  checkAdjust(del('Işık'), [["'IŞIK'!A3+'ışık'!A3+işik!A3+İŞİK!A3+Isik!A3", "'IŞIK'!A1+'ışık'!A1+işik!A1+İŞİK!A1+Isik!A3"]], 'Other');
  // The sheet of the formula is compared the same way.
  checkAdjust(del('IŞIK'), [['A3', 'A1']], 'ışık');
  checkAdjust(del('i\u0307şlem'), [['A3+İŞLEM!A4', 'A1+İŞLEM!A2']], 'İşlem');
  assert.equal(adjustReferenceText('A3 B1', 'ışık', del('Işık')), 'A1');
});

test('refs adjust: a range whose corners both name the same sheet is one range', () => {
  const del = edit('delete', 'row', 4, 1); // row 5
  checkAdjust(
    del,
    [
      // The prefixes stay as written, on both corners.
      ['Sheet1!A1:Sheet1!A5', 'Sheet1!A1:Sheet1!A4'],
      ['Sheet1!A1:Sheet1!A9', 'Sheet1!A1:Sheet1!A8'],
      ['Sheet1!A5:Sheet1!A9', 'Sheet1!A5:Sheet1!A8'],
      ['Sheet1!A5:Sheet1!B5', 'Sheet1!#REF!'],
      ["'Sheet1'!$A$1:sheet1!$A$9", "'Sheet1'!$A$1:sheet1!$A$8"],
      ['SUM(Sheet1!A1:Sheet1!A5)+1', 'SUM(Sheet1!A1:Sheet1!A4)+1'],
      ['Sheet1:Sheet3!A1:Sheet1:Sheet3!A5', 'Sheet1:Sheet3!A1:Sheet1:Sheet3!A4'],
      // Different sheets: two references, as before.
      ['Sheet1!A1:Sheet2!A5', 'Sheet1!A1:Sheet2!A5'],
      ['Sheet1!A5:Sheet2!A5', 'Sheet1!#REF!:Sheet2!A5'],
      ['Sheet2!A1:Sheet2!A5', 'Sheet2!A1:Sheet2!A5'],
      // Only the first two corners pair up.
      ['Sheet1!A1:Sheet1!A5:Sheet1!A9', 'Sheet1!A1:Sheet1!A4:Sheet1!A8'],
      // A second corner that is a range or a name is not paired.
      ['Sheet1!A6:Sheet1!Total', 'Sheet1!A5:Sheet1!Total'],
    ],
    'Sheet2',
  );
  checkAdjust(edit('insert', 'row', 2, 1), [
    ['Sheet1!A1:Sheet1!A5', 'Sheet1!A1:Sheet1!A6'],
    ['Sheet1!A3:Sheet1!A5', 'Sheet1!A4:Sheet1!A6'],
    ['Sheet1!A1:Sheet1!A2', 'Sheet1!A1:Sheet1!A2'],
  ]);
  checkAdjust(edit('delete', 'row', 0, 2, 'İşlem'), [["'İşlem'!A1:işlem!A5", "'İşlem'!A1:işlem!A3"]], 'Other');
  assert.equal(adjustReferenceText('Sheet1!$A$1:Sheet1!$A$5 B5', SHEET, del), 'Sheet1!$A$1:Sheet1!$A$4');
});

test('refs: arguments that are not strings do not throw', () => {
  const e = edit('insert', 'row', 0, 1);
  const bad = (value: unknown) => value as string;
  for (const value of [null, undefined]) {
    assert.equal(shiftFormula(bad(value), 1, 1), '');
    assert.equal(adjustFormula(bad(value), SHEET, e), '');
    assert.equal(adjustReferenceText(bad(value), SHEET, e), '');
  }
  assert.equal(shiftFormula(bad(42), 1, 1), '42');
  assert.equal(adjustFormula('A1', bad(undefined), e), 'A1');
  assert.equal(adjustFormula('A1', SHEET, { ...e, sheetName: bad(null) }), 'A1');
  assert.equal(adjustFormula('A1', SHEET, bad(null) as unknown as StructuralEdit), 'A1');
  assert.equal(adjustReferenceText('A1', SHEET, bad(undefined) as unknown as StructuralEdit), 'A1');
  assert.equal(adjustRange(bad(null) as unknown as { top: number; left: number; bottom: number; right: number }, e), undefined);
});

// ---------------------------------------------------------------------------
// adjustRange

test('refs adjustRange: insert', () => {
  const range = { top: 2, left: 1, bottom: 4, right: 3 }; // B3:D5
  const rows = (at: number, count: number) => adjustRange(range, { kind: 'insert', axis: 'row', at, count });
  const cols = (at: number, count: number) => adjustRange(range, { kind: 'insert', axis: 'col', at, count });
  assert.deepEqual(rows(0, 2), { top: 4, left: 1, bottom: 6, right: 3 });
  assert.deepEqual(rows(2, 2), { top: 4, left: 1, bottom: 6, right: 3 }); // at the start: moves
  assert.deepEqual(rows(3, 2), { top: 2, left: 1, bottom: 6, right: 3 }); // inside: grows
  assert.deepEqual(rows(4, 2), { top: 2, left: 1, bottom: 6, right: 3 }); // before the last row: grows
  assert.deepEqual(rows(5, 2), range); // just after the end
  assert.deepEqual(cols(1, 1), { top: 2, left: 2, bottom: 4, right: 4 });
  assert.deepEqual(cols(2, 1), { top: 2, left: 1, bottom: 4, right: 4 });
  assert.deepEqual(cols(4, 1), range);
  // The input is not modified and a new object comes back.
  assert.notEqual(rows(5, 2), range);
  assert.deepEqual(range, { top: 2, left: 1, bottom: 4, right: 3 });

  const atEnd = { top: MAX_ROWS - 3, left: MAX_COLS - 2, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 };
  assert.deepEqual(adjustRange(atEnd, { kind: 'insert', axis: 'row', at: 0, count: 1 }), { ...atEnd, top: MAX_ROWS - 2 });
  assert.equal(adjustRange(atEnd, { kind: 'insert', axis: 'row', at: 0, count: 3 }), undefined);
  assert.deepEqual(adjustRange(atEnd, { kind: 'insert', axis: 'col', at: 0, count: 1 }), { ...atEnd, left: MAX_COLS - 1 });
  assert.equal(adjustRange(atEnd, { kind: 'insert', axis: 'col', at: 5, count: 2 }), undefined);
});

test('refs adjustRange: delete', () => {
  const range = { top: 2, left: 1, bottom: 4, right: 3 }; // B3:D5
  const rows = (at: number, count: number) => adjustRange(range, { kind: 'delete', axis: 'row', at, count });
  const cols = (at: number, count: number) => adjustRange(range, { kind: 'delete', axis: 'col', at, count });
  assert.deepEqual(rows(0, 2), { top: 0, left: 1, bottom: 2, right: 3 }); // before
  assert.deepEqual(rows(0, 3), { top: 0, left: 1, bottom: 1, right: 3 }); // before and the first row
  assert.deepEqual(rows(2, 1), { top: 2, left: 1, bottom: 3, right: 3 }); // first row
  assert.deepEqual(rows(3, 1), { top: 2, left: 1, bottom: 3, right: 3 }); // middle
  assert.deepEqual(rows(4, 1), { top: 2, left: 1, bottom: 3, right: 3 }); // last row
  assert.deepEqual(rows(4, 10), { top: 2, left: 1, bottom: 3, right: 3 }); // last row and beyond
  assert.deepEqual(rows(5, 3), range); // after
  assert.equal(rows(2, 3), undefined); // exactly all
  assert.equal(rows(0, 10), undefined); // all and more
  assert.deepEqual(cols(0, 1), { top: 2, left: 0, bottom: 4, right: 2 });
  assert.deepEqual(cols(2, 1), { top: 2, left: 1, bottom: 4, right: 2 });
  assert.equal(cols(1, 3), undefined);
  assert.deepEqual(cols(4, 1), range);
  // A single cell.
  const cell = { top: 5, left: 5, bottom: 5, right: 5 };
  assert.equal(adjustRange(cell, { kind: 'delete', axis: 'row', at: 5, count: 1 }), undefined);
  assert.deepEqual(adjustRange(cell, { kind: 'delete', axis: 'row', at: 4, count: 1 }), { top: 4, left: 5, bottom: 4, right: 5 });
  assert.deepEqual(adjustRange(cell, { kind: 'delete', axis: 'row', at: 6, count: 1 }), cell);
  // Corners in the wrong order are put right.
  assert.deepEqual(adjustRange({ top: 4, left: 3, bottom: 2, right: 1 }, { kind: 'delete', axis: 'row', at: 0, count: 1 }), { top: 1, left: 1, bottom: 3, right: 3 });
});

// ---------------------------------------------------------------------------
// adjustReferenceText

test('refs adjustReferenceText: lists and prefixed references', () => {
  const del = edit('delete', 'row', 4, 1); // row 5
  const cases: [string, StructuralEdit, string][] = [
    ['A1:B2 D5', del, 'A1:B2'],
    ['A5 B5', del, ''],
    ['A1 A5 A9', del, 'A1 A8'],
    ['A5 A1  A9', del, 'A1 A8'],
    ['A1:A9', del, 'A1:A8'],
    ['A1:B2 D3', del, 'A1:B2 D3'],
    ['A1:B2  D6', del, 'A1:B2  D5'],
    ['', del, ''],
    ['Sheet1!$A$1:$B$5', edit('delete', 'row', 1, 2), 'Sheet1!$A$1:$B$3'],
    // UNSURE: Excel turns a defined name whose target was deleted into "Sheet1!#REF!"; the
    // list rule (drop the item) is applied here as specified.
    ['Sheet1!$A$5', del, ''],
    ['Sheet2!$A$5', del, 'Sheet2!$A$5'],
    ["'Veri Sayfası'!$A$1:$A$9", edit('insert', 'row', 3, 1, 'Veri Sayfası'), "'Veri Sayfası'!$A$1:$A$10"],
    ['A1:C1 E1', edit('insert', 'col', 1, 1), 'A1:D1 F1'],
    ['$A:$B', edit('delete', 'col', 0, 1), '$A:$A'],
    ['$1:$2 $A:$B', edit('delete', 'row', 0, 2), '$A:$B'],
    [`A${MAX_ROWS} A1`, edit('insert', 'row', 0, 1), 'A2'],
    ['[1]Sheet1!A5 A5 A6', del, '[1]Sheet1!A5 A5'],
    // Anything beyond a plain list is treated as a formula.
    ['SUM(A5)', del, 'SUM(#REF!)'],
    ['Sheet1!$A:$B,Sheet1!$5:$5', del, 'Sheet1!$A:$B,Sheet1!#REF!'],
    ['OFFSET(Sheet1!$A$6,0,0,COUNTA(Sheet1!$A:$A),1)', del, 'OFFSET(Sheet1!$A$5,0,0,COUNTA(Sheet1!$A:$A),1)'],
    ['=A5', del, '=#REF!'],
    ['Total A5', del, 'Total #REF!'],
  ];
  for (const [input, e, expected] of cases) {
    assert.equal(adjustReferenceText(input, SHEET, e), expected, `${e.kind} ${e.axis} ${e.at}+${e.count}: ${JSON.stringify(input)}`);
  }
  // The default sheet decides for references without a prefix.
  assert.equal(adjustReferenceText('A5 B9', 'Sheet2', del), 'A5 B9');
});

// ---------------------------------------------------------------------------
// Properties over generated formulas

/** Small deterministic generator (mulberry32). */
function random(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function randomFormula(rnd: () => number, lowerCase: boolean, dollars = 0.3): string {
  const pick = <T>(list: T[]): T => list[Math.floor(rnd() * list.length)];
  const int = (max: number) => Math.floor(rnd() * max);
  const dollar = () => (rnd() < dollars ? '$' : '');
  const letters = () => {
    const name = pick(['A', 'B', 'C', 'Z', 'AA', 'AZ', 'BC']);
    return lowerCase && rnd() < 0.3 ? name.toLowerCase() : name;
  };
  const cell = () => `${dollar()}${letters()}${dollar()}${1 + int(200)}`;
  const reference = () => {
    const kind = int(10);
    const prefix = pick(['', '', '', 'Sheet1!', 'Sheet2!', "'Veri Sayfası'!", "'It''s'!", '[1]Sheet1!', 'Sheet1:Sheet3!']);
    if (kind < 5) return prefix + cell();
    if (kind < 8) return `${prefix}${cell()}:${cell()}`;
    if (kind < 9) return `${prefix}${dollar()}${letters()}:${dollar()}${letters()}`;
    return `${prefix}${dollar()}${1 + int(200)}:${dollar()}${1 + int(200)}`;
  };
  const filler = () => pick(['"A1:B2"', '"say ""C3"""', '1E5', '2.5E-3', '42', 'Toplam', 'A1B2C', '_A1', '#REF!', '#N/A', '#DIV/0!', 'Tablo1[Sütun]', '[@Sütun]', 'TRUE']);
  const operator = () => pick(['+', '-', '*', '/', '&', ' ', ',', ' + ', '<>', '^']);
  const term = (depth: number): string => {
    const kind = int(10);
    if (kind < 5) return reference();
    if (kind < 7) return filler();
    if (depth > 2) return reference();
    const name = pick(['SUM', 'IF', 'LOG10', 'ATAN2', 'DAYS360', 'A1B', 'vlookup', '']);
    const args: string[] = [];
    for (let k = 1 + int(3); k > 0; k--) args.push(term(depth + 1));
    return `${name}(${args.join(pick([',', ', ', ';']))})`;
  };
  let formula = term(0);
  for (let k = int(5); k > 0; k--) formula += operator() + term(0);
  return formula;
}

test('refs property: untouched formulas come back identical', () => {
  const rnd = random(1);
  for (let k = 0; k < 2000; k++) {
    const formula = randomFormula(rnd, true);
    assert.equal(shiftFormula(formula, 0, 0), formula);
    // The generated references stay below row 201 and column BD.
    assert.equal(adjustFormula(formula, SHEET, edit('insert', 'row', 500, 3)), formula);
    assert.equal(adjustFormula(formula, SHEET, edit('delete', 'col', 500, 3)), formula);
    // An edit on a sheet no reference names.
    assert.equal(adjustFormula(formula, 'Sheet9', edit('delete', 'row', 0, 50, 'Other')), formula);
    assert.equal(adjustFormula(formula, 'Sheet9', edit('insert', 'col', 0, 50, 'Other')), formula);
  }
});

test('refs property: shifting there and back', () => {
  const rnd = random(2);
  for (let k = 0; k < 2000; k++) {
    // Upper case only: a rewritten reference is written in upper case.
    const formula = randomFormula(rnd, false);
    const dRow = Math.floor(rnd() * 50);
    const dCol = Math.floor(rnd() * 50);
    // Moving away from the top-left corner never leaves the sheet, so nothing turns into #REF!.
    assert.equal(shiftFormula(shiftFormula(formula, dRow, dCol), -dRow, -dCol), formula);
  }
});

test('refs property: an insertion followed by the same deletion is undone', () => {
  // This holds because the generated references are far from the last row and column and are
  // written in upper case. It does not hold when the insertion pushed something off the sheet
  // (#REF! or a clipped range), and never in the other order: deleting first loses the deleted
  // cells (#REF!) and the shrunk ranges, which the insertion grows back only when it falls
  // strictly inside them.
  const rnd = random(3);
  for (let k = 0; k < 2000; k++) {
    const formula = randomFormula(rnd, false);
    const axis = rnd() < 0.5 ? 'row' : 'col';
    const at = Math.floor(rnd() * (axis === 'row' ? 210 : 60));
    const count = 1 + Math.floor(rnd() * 5);
    const inserted = adjustFormula(formula, SHEET, edit('insert', axis, at, count));
    assert.equal(adjustFormula(inserted, SHEET, edit('delete', axis, at, count)), formula, `${axis} ${at}+${count}: ${formula}`);
  }
  // The other order, spelled out.
  const deleted = adjustFormula('A1:A3+A2', SHEET, edit('delete', 'row', 1, 1));
  assert.equal(deleted, 'A1:A2+#REF!');
  assert.equal(adjustFormula(deleted, SHEET, edit('insert', 'row', 1, 1)), 'A1:A3+#REF!');
  const shrunk = adjustFormula('A1:A2', SHEET, edit('delete', 'row', 1, 1));
  assert.equal(adjustFormula(shrunk, SHEET, edit('insert', 'row', 1, 1)), 'A1:A1');
});

// ---------------------------------------------------------------------------
// adjustAnchoredFormula

/** Rows of [input, expected] for one edit and one move of the anchor cell. */
function checkAnchored(e: StructuralEdit, before: [number, number], after: [number, number], cases: [string, string][], formulaSheet = SHEET): void {
  const b = { row: before[0], col: before[1] };
  const a = { row: after[0], col: after[1] };
  for (const [input, expected] of cases) {
    assert.equal(adjustAnchoredFormula(input, formulaSheet, e, b, a), expected, `${e.kind} ${e.axis} ${e.at}+${e.count}, anchor ${before} -> ${after}: ${JSON.stringify(input)}`);
  }
}

test('refs anchored: deleting rows', () => {
  // sqref A5:A10, row 5 deleted: the rule still starts at A5 (the anchor is inside the deleted rows).
  checkAnchored(edit('delete', 'row', 4, 1), [4, 0], [4, 0], [
    ['A5>B5', 'A5>B5'],
    ['A7', 'A7'],
    ['$A$7', '$A$6'],
    ['A$7', 'A$6'],
    ['$A7', '$A7'],
    ['$A$5', '#REF!'],
    ['A$5', '#REF!'],
    ['$A$4', '$A$4'],
    ['AND(A5>$B$7,"A5"<>a5)', 'AND(A5>$B$6,"A5"<>a5)'],
    ['$A$1:$A$10', '$A$1:$A$9'],
    ['$A$5:$A$5', '#REF!'],
    ['SUM($A$5:A5)', 'SUM(#REF!)'],
    ['SUM($A$1:A5)', 'SUM($A$1:A5)'],
    ['SUM($A$6:A5)', 'SUM($A$5:A5)'],
    ['$5:$5', '#REF!'],
    ['$3:$9', '$3:$8'],
    ['5:9', '5:9'],
    ['$A:$C', '$A:$C'],
  ]);
  // Anchor below the deleted rows: sqref A5:A10 -> A3:A8 after rows 1-2 go.
  checkAnchored(edit('delete', 'row', 0, 2), [4, 0], [2, 0], [
    ['A5>B5', 'A3>B3'],
    ['$A$7', '$A$5'],
    ['$A$2', '#REF!'],
    ['A1', '#REF!'], // two rows above the sheet
    ['A2:A5', '#REF!'],
    ['$A$3:A5', '$A$1:A3'],
    ['A$9+$A9', 'A$7+$A7'],
  ]);
  // Anchor above the deleted rows: it stays.
  checkAnchored(edit('delete', 'row', 20, 2), [4, 0], [4, 0], [
    ['A5>B5', 'A5>B5'],
    ['A30', 'A30'],
    ['$A$30', '$A$28'],
    ['$A$21', '#REF!'],
    ['$A$5:$A$30', '$A$5:$A$28'],
  ]);
});

test('refs anchored: inserting rows', () => {
  // Three rows inserted before row 3: sqref A5:A10 -> A8:A13.
  checkAnchored(edit('insert', 'row', 2, 3), [4, 0], [7, 0], [
    ['A5>B5', 'A8>B8'],
    ['$C$1', '$C$1'],
    ['$C$3', '$C$6'],
    ['A$3', 'A$6'],
    ['$A5', '$A8'],
    ['a5', 'A8'],
    ['$A$1:$A$10', '$A$1:$A$13'],
    ['$A$3:$A$10', '$A$6:$A$13'],
    ['$A$1:A5', '$A$1:A8'],
    ['$A$3:A5', '$A$6:A8'],
    ['A1:A2', 'A4:A5'],
    ['$1:$3', '$1:$6'],
    ['1:3', '4:6'],
    ['A:A', 'A:A'],
    [`A${MAX_ROWS - 2}`, '#REF!'],
    [`$A$${MAX_ROWS}`, '#REF!'],
  ]);
  // Inserted inside the rule: the anchor stays, anchored targets below move.
  checkAnchored(edit('insert', 'row', 6, 1), [4, 0], [4, 0], [
    ['A5>B5', 'A5>B5'],
    ['A9', 'A9'],
    ['$A$9', '$A$10'],
    ['$A$5:$A$10', '$A$5:$A$11'],
  ]);
});

test('refs anchored: columns', () => {
  // Two columns inserted before B: sqref C1:C9 -> E1:E9.
  checkAnchored(edit('insert', 'col', 1, 2), [0, 2], [0, 4], [
    ['C1>$B1', 'E1>$D1'],
    ['$A1', '$A1'],
    ['C$1', 'E$1'],
    ['$B$1:$C$5', '$D$1:$E$5'],
    ['$A1:$C1', '$A1:$E1'],
    ['$A1:C1', '$A1:E1'],
    ['B:B', 'D:D'],
    ['$B:$B', '$D:$D'],
    ['$A:$B', '$A:$D'],
    ['1:1', '1:1'],
    ['$1:$1', '$1:$1'],
    ['XFC1', '#REF!'],
  ]);
  // Column C deleted out of sqref B1:C5: the anchor stays on B.
  checkAnchored(edit('delete', 'col', 2, 1), [0, 1], [0, 1], [
    ['B1>C1', 'B1>C1'],
    ['$C1', '#REF!'],
    ['$D1', '$C1'],
    ['$B1', '$B1'],
    ['$B$1:$D$1', '$B$1:$C$1'],
    ['$C:$C', '#REF!'],
    ['C:C', 'C:C'],
  ]);
  // Column A deleted: sqref B1:C5 -> A1:B5.
  checkAnchored(edit('delete', 'col', 0, 1), [0, 1], [0, 0], [
    ['B1>C1', 'A1>B1'],
    ['A1', '#REF!'],
    ['$A1', '#REF!'],
    ['$B1', '$A1'],
    ['B$1:$D$9', 'A$1:$C$9'],
  ]);
});

test('refs anchored: other sheets and other workbooks', () => {
  const e = edit('insert', 'row', 2, 3);
  // Relative parts follow the anchor whatever they point at; anchored parts follow the edit of their own sheet only.
  checkAnchored(e, [4, 0], [7, 0], [
    ['Sheet2!A5', 'Sheet2!A8'],
    ['Sheet2!$A$5', 'Sheet2!$A$5'],
    ['Sheet2!$A5:$B$9', 'Sheet2!$A8:$B$9'],
    ['Sheet1!$A$5+sheet1!A5', 'Sheet1!$A$8+sheet1!A8'],
    ["'Veri Sayfası'!A5>Özet!$B$5", "'Veri Sayfası'!A8>Özet!$B$5"],
    ['[1]Sheet1!A5+[1]Sheet1!$A$5', '[1]Sheet1!A8+[1]Sheet1!$A$5'],
    ['Sheet1:Sheet3!$A$5', 'Sheet1:Sheet3!$A$8'],
    ['Sheet2:Sheet3!$A$5', 'Sheet2:Sheet3!$A$5'],
    ['Sheet1!$A$1:Sheet1!$A$5', 'Sheet1!$A$1:Sheet1!$A$8'],
    ['Tablo1[A5]+LOG10(A5)+TAX2024+"A5"', 'Tablo1[A5]+LOG10(A8)+TAX2027+"A5"'],
  ]);
  // A formula on another sheet than the edited one: bare references are not about the edit.
  checkAnchored(
    e,
    [4, 0],
    [7, 0],
    [['A5+$A$5+Sheet1!$A$5', 'A8+$A$5+Sheet1!$A$8']],
    'Sheet2',
  );
});

test('refs anchored: nothing to do', () => {
  const formula = 'and(a5>b5, Sheet2!c$1="x")';
  const anchor = { row: 4, col: 0 };
  // No anchored part on the edited sheet and an anchor that stayed: identical.
  assert.equal(adjustAnchoredFormula(formula, SHEET, edit('delete', 'row', 4, 1), anchor, anchor), formula);
  assert.equal(adjustAnchoredFormula('A5>B5', SHEET, edit('delete', 'row', 0, 100), anchor, anchor), 'A5>B5');
  // An edit that moves nothing still lets the anchor move.
  assert.equal(adjustAnchoredFormula('A5>$B$5', SHEET, edit('insert', 'row', 0, 0), anchor, { row: 5, col: 1 }), 'B6>$B$5');
  assert.equal(adjustAnchoredFormula('A5>$B$5', SHEET, edit('insert', 'row', 0, 0), anchor, anchor), 'A5>$B$5');
  const bad = (value: unknown) => value as string;
  assert.equal(adjustAnchoredFormula(bad(null), SHEET, edit('insert', 'row', 0, 1), anchor, anchor), '');
  assert.equal(adjustAnchoredFormula('A5+$A$5', SHEET, edit('insert', 'row', 0, 1), bad(null) as unknown as typeof anchor, anchor), 'A5+$A$6');
  assert.equal(adjustAnchoredFormula('A5', SHEET, bad(undefined) as unknown as StructuralEdit, anchor, anchor), 'A5');
});

test('refs property: anchored formulas agree with adjustFormula and shiftFormula', () => {
  const rnd = random(5);
  for (let k = 0; k < 2000; k++) {
    const axis = rnd() < 0.5 ? 'row' : 'col';
    const e = edit(rnd() < 0.5 ? 'insert' : 'delete', axis, Math.floor(rnd() * (axis === 'row' ? 210 : 60)), 1 + Math.floor(rnd() * 5));
    const before = { row: Math.floor(rnd() * 20), col: Math.floor(rnd() * 20) };
    const after = { row: Math.floor(rnd() * 20), col: Math.floor(rnd() * 20) };
    // Every part anchored: the anchor cell does not matter.
    const anchored = randomFormula(rnd, true, 1);
    assert.equal(adjustAnchoredFormula(anchored, SHEET, e, before, after), adjustFormula(anchored, SHEET, e), anchored);
    // No part anchored: the edit does not matter, only the move of the anchor.
    const relative = randomFormula(rnd, true, 0);
    assert.equal(adjustAnchoredFormula(relative, SHEET, e, before, after), shiftFormula(relative, after.row - before.row, after.col - before.col), relative);
    // An anchor that stayed and no anchored part: identical.
    assert.equal(adjustAnchoredFormula(relative, SHEET, e, before, before), relative);
  }
});

// ---------------------------------------------------------------------------
// Robustness

test('refs: malformed input never throws and keeps what it cannot read', () => {
  const e = edit('insert', 'row', 0, 1);
  const cases: [string, string][] = [
    // From an unclosed quote or bracket on, nothing is touched.
    ['A1+"abc+B1', 'A2+"abc+B1'],
    ["A1+'abc+B1", "A2+'abc+B1"],
    ['A1+[abc+B1', 'A2+[abc+B1'],
    ['A1+Tablo1[[x]+B1', 'A2+Tablo1[[x]+B1'],
    // Unbalanced parentheses do not matter to the scanner.
    ['SUM(A1', 'SUM(A2'],
    ['A1)))+((B1', 'A2)))+((B2'],
    // A quoted name that is not a sheet prefix.
    ["'abc'+A1", "'abc'+A2"],
    ['A1:', 'A2:'],
    [':A1', ':A2'],
    ['A1::B1', 'A2::B2'],
    ['A1:B1:C1', 'A2:B2:C2'],
    ['!A1', '!A2'],
    ['Sheet1!', 'Sheet1!'],
    ['Sheet1!!A1', 'Sheet1!!A2'],
    ['$', '$'],
    ['$$A1', '$$A1'],
    ['A$', 'A$'],
    ['$A$', '$A$'],
    ['A1$', 'A1$'],
    ['#', '#'],
    ['#REF', '#REF'],
    ['A1]', 'A2]'],
    ['\u0000A1\uffff', '\u0000A1\uffff'],
    ['\ud83d\ude00+A1', '\ud83d\ude00+A2'],
    ['99999999999999999999:1', '99999999999999999999:1'],
    ['A99999999999999999999', 'A99999999999999999999'],
  ];
  for (const [input, expected] of cases) {
    assert.equal(adjustFormula(input, SHEET, e), expected, JSON.stringify(input));
    assert.equal(shiftFormula(input, 1, 0), expected, JSON.stringify(input));
    assert.doesNotThrow(() => adjustReferenceText(input, SHEET, e));
    assert.equal(shiftFormula(input, 0, 0), input);
  }
});

test('refs: random garbage never throws', () => {
  const rnd = random(4);
  const alphabet = 'A1$:!\'"[]()#, .BZ9e+-xfd/\\?_Ş';
  for (let k = 0; k < 5000; k++) {
    let text = '';
    for (let len = Math.floor(rnd() * 24); len > 0; len--) text += alphabet[Math.floor(rnd() * alphabet.length)];
    assert.equal(typeof shiftFormula(text, 1, -1), 'string');
    assert.equal(typeof adjustFormula(text, SHEET, edit('delete', 'col', 1, 2)), 'string');
    assert.equal(typeof adjustReferenceText(text, SHEET, edit('insert', 'row', 1, 2)), 'string');
    assert.equal(shiftFormula(text, 0, 0), text);
  }
});

/** Best of a few runs, so a busy machine does not fail the test. */
function fastest(run: () => void): number {
  let best = Infinity;
  for (let k = 0; k < 5; k++) {
    const start = performance.now();
    run();
    best = Math.min(best, performance.now() - start);
  }
  return best;
}

test('refs: 100,000 characters in linear time', () => {
  const size = 100000;
  const fill = (unit: string) => unit.repeat(Math.ceil(size / unit.length)).slice(0, size);
  const inputs = [
    fill("SUM(A1:B2,Sheet1!$C$3,'Veri Sayfası'!D4)+"),
    fill('A1+'),
    fill('a!A1:'),
    fill("'a'!A1:b!A1:"),
    fill('A1:'),
    fill('A'),
    fill('A:'),
    fill('1'),
    fill("'"),
    fill("'a'"),
    fill("'a'!"),
    fill('"'),
    fill('['),
    fill('[a]'),
    fill(']'),
    fill('('),
    fill('!'),
    fill('a!'),
    fill('#'),
    fill('#REF!'),
    fill('$'),
    fill('$A$1:'),
    fill("A1 'x"),
    fill('A1 "x'),
    fill('A1 [x'),
    fill('A1 '),
  ];
  const e = edit('insert', 'row', 0, 1);
  for (const input of inputs) {
    const label = JSON.stringify(input.slice(0, 12));
    assert.ok(fastest(() => shiftFormula(input, 1, 1)) < 1000, `shiftFormula ${label}`);
    assert.ok(fastest(() => adjustFormula(input, SHEET, e)) < 1000, `adjustFormula ${label}`);
    assert.ok(fastest(() => adjustReferenceText(input, SHEET, e)) < 1000, `adjustReferenceText ${label}`);
  }
  const many = fill('A1+');
  assert.equal(shiftFormula(many, 1, 0), fill('A2+'));
});

// ---------------------------------------------------------------------------
// Whole columns and rows, and references that go round the edge of the sheet

test('refs adjust: a span over a whole axis stays whole', () => {
  const lastRow = MAX_ROWS;
  // Whole columns: rows inserted or deleted change nothing, columns move them.
  for (const e of [edit('insert', 'row', 0, 1), edit('insert', 'row', 7, 3), edit('delete', 'row', 0, 5), edit('delete', 'row', lastRow - 1, 1)]) {
    checkAdjust(e, [
      ['B:B', 'B:B'],
      [`B1:B${lastRow}`, `B1:B${lastRow}`],
      [`$B$1:$C$${lastRow}`, `$B$1:$C$${lastRow}`],
      [`SUM(B${lastRow}:B1)`, `SUM(B${lastRow}:B1)`],
      [`1:${lastRow}`, `1:${lastRow}`],
    ]);
    assert.deepEqual(adjustRange({ top: 0, left: 1, bottom: MAX_ROWS - 1, right: 1 }, e), { top: 0, left: 1, bottom: MAX_ROWS - 1, right: 1 });
    assert.ok(adjustReferenceText(`D9:D12 B:B B1:B${lastRow}`, SHEET, e).endsWith(` B:B B1:B${lastRow}`));
  }
  checkAdjust(edit('insert', 'col', 0, 1), [
    ['B:B', 'C:C'],
    [`B1:B${lastRow}`, `C1:C${lastRow}`],
    ['A3:XFD3', 'A3:XFD3'],
    ['3:3', '3:3'],
    ['A:XFD', 'A:XFD'],
  ]);
  checkAdjust(edit('delete', 'col', 1, 1), [
    ['B:C', 'B:B'],
    [`B1:C${lastRow}`, `B1:B${lastRow}`],
    ['B:B', '#REF!'],
    ['A3:XFD3', 'A3:XFD3'],
  ]);
  // Almost whole is not whole.
  checkAdjust(edit('insert', 'row', 0, 1), [
    [`B2:B${lastRow}`, `B3:B${lastRow}`],
    [`B1:B${lastRow - 1}`, `B2:B${lastRow}`],
    ['3:3', '4:4'],
  ]);
  assert.deepEqual(adjustRange({ top: 2, left: 0, bottom: 2, right: MAX_COLS - 1 }, edit('insert', 'col', 3, 2)), { top: 2, left: 0, bottom: 2, right: MAX_COLS - 1 });
  assert.deepEqual(adjustRange({ top: 2, left: 0, bottom: 2, right: MAX_COLS - 1 }, edit('insert', 'row', 0, 1)), { top: 3, left: 0, bottom: 3, right: MAX_COLS - 1 });
  // Lists keep the form each reference was written in.
  assert.equal(adjustReferenceText('D1:D4 B:B', SHEET, edit('insert', 'col', 0, 1)), 'E1:E4 C:C');
  assert.equal(adjustReferenceText('3:3 A1', SHEET, edit('insert', 'row', 0, 1)), '4:4 A2');
  assert.equal(adjustReferenceText('B:B D1', SHEET, edit('delete', 'col', 1, 1)), 'C1');
  // Anchored formulas: a fully anchored whole column behaves the same.
  assert.equal(adjustAnchoredFormula(`SUM($B$1:$B$${lastRow})+B1`, SHEET, edit('insert', 'row', 0, 1), { row: 0, col: 1 }, { row: 0, col: 1 }), `SUM($B$1:$B$${lastRow})+B1`);
});

test('refs shift: with wrap, a reference pushed off the sheet comes back at the other edge', () => {
  const cases: [string, number, number, string][] = [
    ['A1', -1, 0, `A${MAX_ROWS}`],
    ['A1', 0, -1, 'XFD1'],
    ['A1', -2, -2, `XFC${MAX_ROWS - 1}`],
    [`A${MAX_ROWS}`, 1, 0, 'A1'],
    ['XFD1', 0, 2, 'B1'],
    ['$A1', -1, -1, `$A${MAX_ROWS}`],
    ['A$1', -1, -1, 'XFD$1'],
    ['$A$1', -1, -1, '$A$1'],
    ['A1:B2', -1, 0, `A${MAX_ROWS}:B1`],
    ['A:B', 0, -1, 'XFD:A'],
    ['1:2', -1, 0, `${MAX_ROWS}:1`],
    ['Sheet2!A1+B5', -1, 0, `Sheet2!A${MAX_ROWS}+B4`],
    ['B5', 2, 3, 'E7'],
    ['"A1"&A1', -1, 0, `"A1"&A${MAX_ROWS}`],
  ];
  for (const [input, dRow, dCol, expected] of cases) assert.equal(shiftFormula(input, dRow, dCol, true), expected, input);
  // Without the flag nothing changed.
  assert.equal(shiftFormula('A1', -1, 0), '#REF!');
  assert.equal(shiftFormula('A1', -1, 0, false), '#REF!');
});

test('refs adjust: anchored formulas with wrap', () => {
  const none = edit('delete', 'row', 0, 1);
  // The anchor moved up one row: "the cell above" of row 1 is the last row.
  assert.equal(adjustAnchoredFormula('B1>0', SHEET, none, { row: 1, col: 1 }, { row: 0, col: 1 }, true), `B${MAX_ROWS}>0`);
  assert.equal(adjustAnchoredFormula('B1>0', SHEET, none, { row: 1, col: 1 }, { row: 0, col: 1 }), '#REF!>0');
  assert.equal(adjustAnchoredFormula('A2>0', SHEET, edit('delete', 'col', 0, 1), { row: 1, col: 1 }, { row: 1, col: 0 }, true), 'XFD2>0');
  assert.equal(adjustAnchoredFormula(`A${MAX_ROWS}>0`, SHEET, edit('insert', 'row', 0, 1), { row: 0, col: 0 }, { row: 1, col: 0 }, true), 'A1>0');
  // What is anchored with "$" names a cell, and a deleted cell is gone whatever the flag says.
  assert.equal(adjustAnchoredFormula('$B$1>B1', SHEET, none, { row: 1, col: 1 }, { row: 0, col: 1 }, true), `#REF!>B${MAX_ROWS}`);
  assert.equal(adjustAnchoredFormula('$B1>0', SHEET, edit('delete', 'col', 1, 1), { row: 0, col: 2 }, { row: 0, col: 1 }, true), '#REF!>0');
});

test('refs: how far the relative references of a formula reach', () => {
  assert.deepEqual(relativeExtent('B3+D10'), { minRow: 2, maxRow: 9, minCol: 1, maxCol: 3 });
  assert.deepEqual(relativeExtent('$B3+D$10'), { minRow: 2, maxRow: 2, minCol: 3, maxCol: 3 });
  assert.deepEqual(relativeExtent('$B$3+"C7"+1'), { minRow: Infinity, maxRow: -1, minCol: Infinity, maxCol: -1 });
  assert.deepEqual(relativeExtent('SUM(C:D)+SUM(5:6)'), { minRow: 4, maxRow: 5, minCol: 2, maxCol: 3 });
  assert.deepEqual(relativeExtent('SUM(B2:$E$9)+Other!A1'), { minRow: 0, maxRow: 1, minCol: 0, maxCol: 1 });
  assert.deepEqual(relativeExtent(undefined as unknown as string), { minRow: Infinity, maxRow: -1, minCol: Infinity, maxCol: -1 });
  // A megabyte of text is one pass.
  const start = Date.now();
  assert.equal(relativeExtent('A1+'.repeat(350000)).maxRow, 0);
  assert.ok(Date.now() - start < 10000);
});
