// Formulas typed the Turkish way (";" between arguments, "1,5", Turkish
// function names) and stored the one way files have them, and back.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { parseFormula } from '../src/calc/parse';
import { formulaFromLocal, formulaToLocal } from '../src/edit/formula-locale';
import { cellEditText, parseCellInput } from '../src/edit/input';
import { LOCALE_EN, LOCALE_TR } from '../src/format/locale';
import { createSheet, emptyStyleSheet, getCell, type Cell, type Workbook } from '../src/model/sheet/types';
import { copyRange, editTextAt, pasteAt, setCellInput } from '../src/ops/commands';
import { applyOp } from '../src/ops/ops';

const TR = LOCALE_TR;

function workbook(): Workbook {
  const sheet = createSheet({ name: 'Data', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  return { sheets: [sheet], styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 };
}

// [stored, typed in Turkish]: each has to give the other.
const PAIRS: [string, string][] = [
  ['SUM(A1,B1)', 'SUM(A1;B1)'],
  ['A1*1.5', 'A1*1,5'],
  ['IF(A1>2.5,"a,b;c",1)', 'IF(A1>2,5;"a,b;c";1)'],
  ['ROUND(1.2345,2)+0.5', 'ROUND(1,2345;2)+0,5'],
  ["'A;B,1.5'!A1+SUM('It''s, 1.5'!B2:C3,1)", "'A;B,1.5'!A1+SUM('It''s, 1.5'!B2:C3;1)"],
  ['SUM((A1:A2,C1:C2))', 'SUM((A1:A2;C1:C2))'],
  ['1.5E+3+2E-2', '1,5E+3+2E-2'],
  ['SUM(1:1,3:3)', 'SUM(1:1;3:3)'],
  ['{1,2;3,4}', '{1;2\\3;4}'],
  ['SUM({1.5,2;"a;b",TRUE})', 'SUM({1,5;2\\"a;b";TRUE})'],
  ['Table1[[#This Row],[Price, net]]*1.2', 'Table1[[#This Row],[Price, net]]*1,2'],
  ['[1]Sheet1!A1+[2]Sheet2!B2', '[1]Sheet1!A1+[2]Sheet2!B2'],
  ['IFERROR(A1/B1,#N/A)', 'IFERROR(A1/B1;#N/A)'],
  ['A1&"x"&B1', 'A1&"x"&B1'],
  ['my.name1+Sheet.2!A1', 'my.name1+Sheet.2!A1'],
  ['SUM($A$1:$B$2,Data!C3)%', 'SUM($A$1:$B$2;Data!C3)%'],
  ['"unterminated, 1.5', '"unterminated, 1.5'],
  ['T.TEST(A1:A5,B1:B5,2,1)', 'T.TEST(A1:A5;B1:B5;2;1)'],
];

test('formula locale: stored and Turkish notation translate into each other', () => {
  for (const [stored, typed] of PAIRS) {
    assert.equal(formulaToLocal(stored, TR), typed, `to local: ${stored}`);
    assert.equal(formulaFromLocal(typed, TR), stored, `from local: ${typed}`);
  }
});

test('formula locale: English and missing locales change nothing', () => {
  for (const [stored, typed] of PAIRS) {
    assert.equal(formulaToLocal(stored, LOCALE_EN), stored);
    assert.equal(formulaFromLocal(typed, LOCALE_EN), typed);
    assert.equal(formulaToLocal(stored), stored);
    assert.equal(formulaFromLocal(stored, undefined), stored);
  }
});

test('formula locale: numbers that have no Turkish spelling of their own are written so that they read back', () => {
  // ".5" would read as a separator and a 5; "1." as a 1 and a separator.
  assert.equal(formulaToLocal('.5+A1', TR), '0,5+A1');
  assert.equal(formulaFromLocal('0,5+A1', TR), '0.5+A1');
  assert.equal(formulaToLocal('SUM(1.,2)', TR), 'SUM(1;2)');
  assert.equal(formulaToLocal('1.E3', TR), '1E3');
});

test('formula locale: a comma is the decimal comma or nothing - a formula typed the English way is refused, never read as another one', () => {
  // Each of these would otherwise turn into a different formula without a word: "5,1" is a number.
  const english = [
    'IF(C2>5,1,0)',
    'VLOOKUP(20,C2:C7,1,0)',
    'DATE(2024,1,5)',
    'MID(A1,2,3)',
    'ROUND(A1,2)',
    'SUM(A1,B1)',
    'SUM(A1,5)',
    'IF(A1>1, 2, 3)',
    'SUM(1,2,3)',
    'SUM(2,A1)',
    '{1, 2}',
    'SUM((A1:A2,C1:C2))',
    '1,',
    ',1',
    'A1,',
  ];
  for (const typed of english) {
    assert.throws(() => formulaFromLocal(typed, TR), /^RangeError: The formula cannot be read$/, typed);
    // The same text is a formula as it stands where the comma separates arguments.
    assert.equal(formulaFromLocal(typed, LOCALE_EN), typed);
  }
  // As in the Turkish Excel: between two digits the comma is the decimal comma, so this is the one number 1,5.
  assert.equal(formulaFromLocal('SUM(1,5)', TR), 'SUM(1.5)');
  assert.equal(formulaFromLocal('SUM(1,5;2,25)', TR), 'SUM(1.5,2.25)');
  assert.equal(formulaFromLocal('{1,5;2\\3;4}', TR), '{1.5,2;3,4}');
  // A point in a number is left to the parser, as before.
  assert.equal(formulaFromLocal('A1*1.5', TR), 'A1*1.5');
  // Commas inside text, sheet names and brackets are not the formula's.
  assert.equal(formulaFromLocal('"a,b"&\'x, y\'!A1&T[[#This Row],[a, b]]', TR), '"a,b"&\'x, y\'!A1&T[[#This Row],[a, b]]');
  // Any locale with a decimal comma does the same.
  const german = { ...LOCALE_EN, decimal: ',', group: '.' };
  assert.throws(() => formulaFromLocal('ROUND(A1,2)', german), /The formula cannot be read/);
});

test('formula locale: what is typed with a stray comma is refused by the commands, and pasted from elsewhere it stays text', () => {
  const wb = workbook();
  const sheet = wb.sheets[0];
  for (const typed of ['=IF(C2>5,1,0)', '=VLOOKUP(20,C2:C7,1,0)', '=DATE(2024,1,5)', '=MID(A1,2,3)', '=ROUND(A1,2)']) {
    assert.throws(() => setCellInput(wb, 0, 0, 0, typed, { locale: TR }), /^RangeError: The formula cannot be read$/, typed);
    assert.equal(parseCellInput(typed, { locale: TR }).invalid, true, typed);
    // The same text in English is the formula it says.
    assert.equal(setCellInput(wb, 0, 0, 0, typed, { locale: LOCALE_EN }).length, 1);
  }
  assert.equal(getCell(sheet, 0, 0), undefined);
  assert.equal(parseCellInput('=ROUND(A1;2)', { locale: TR }).invalid, undefined);
  const result = pasteAt(wb, 0, { top: 0, left: 0, bottom: 0, right: 0 }, '=IF(C2>5,1,0)\t=IF(C2>5;1;0)', undefined, { locale: TR });
  for (const op of result.ops) applyOp(wb, op);
  assert.deepEqual(getCell(sheet, 0, 0), { kind: 'string', value: '=IF(C2>5,1,0)', style: 0 });
  assert.equal(getCell(sheet, 0, 1)!.formula!.text, 'IF(C2>5,1,0)');
  // A copy made here carries the stored formula and is not read again.
  const copied = copyRange(wb, 0, { top: 0, left: 1, bottom: 0, right: 1 }, { locale: TR });
  const pasted = pasteAt(wb, 0, { top: 1, left: 1, bottom: 1, right: 1 }, copied.text, copied, { locale: TR });
  for (const op of pasted.ops) applyOp(wb, op);
  assert.equal(getCell(sheet, 1, 1)!.formula!.text, 'IF(C3>5,1,0)');
});

test('formula locale: DOĞRU and YANLIŞ are constants only where a constant can stand', () => {
  assert.equal(formulaFromLocal('EĞER(A1;DOĞRU;YANLIŞ)', TR), 'IF(A1,TRUE,FALSE)');
  assert.equal(formulaFromLocal('doğru+yanlış', TR), 'TRUE+FALSE');
  assert.equal(formulaFromLocal('DOĞRU()', TR), 'TRUE()');
  // The name of a sheet, of a run of sheets, and a name of a sheet's own.
  assert.equal(formulaFromLocal('DOĞRU!A1+1', TR), 'DOĞRU!A1+1');
  assert.equal(formulaFromLocal('TOPLA(DOĞRU:YANLIŞ!A1)', TR), 'SUM(DOĞRU:YANLIŞ!A1)');
  assert.equal(formulaFromLocal('Sayfa1!DOĞRU*2', TR), 'Sayfa1!DOĞRU*2');
  assert.equal(formulaFromLocal("'DOĞRU'!A1&YANLIŞ", TR), "'DOĞRU'!A1&FALSE");
  // A range of two cells is not a run of sheets.
  assert.equal(formulaFromLocal('EĞER(DOĞRU;A1:B2;0)', TR), 'IF(TRUE,A1:B2,0)');
  // A defined name of the workbook, however it is cased.
  const isName = (name: string) => name === 'DOĞRU' || name === 'doğru';
  assert.equal(formulaFromLocal('DOĞRU*2+YANLIŞ', TR, { isName }), 'DOĞRU*2+FALSE');
  assert.equal(formulaFromLocal('doğru*2', TR, { isName }), 'doğru*2');
  // Function names were always read only where they are called.
  assert.equal(formulaFromLocal('TOPLA!A1+EĞER+TOPLA(EĞER!B1:B3)', TR), 'TOPLA!A1+EĞER+SUM(EĞER!B1:B3)');
  // What is shown never turns English words into Turkish ones.
  assert.equal(formulaToLocal('TRUE!A1+TRUE', TR), 'TRUE!A1+TRUE');

  // Through the commands: a sheet and a defined name called DOĞRU keep their formulas.
  const wb = workbook();
  wb.sheets.push(createSheet({ name: 'DOĞRU', sheetId: 2, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet2.xml' }));
  for (const op of setCellInput(wb, 0, 0, 0, '=DOĞRU!A1+1', { locale: TR })) applyOp(wb, op);
  assert.equal(getCell(wb.sheets[0], 0, 0)!.formula!.text, 'DOĞRU!A1+1');
  for (const op of setCellInput(wb, 0, 1, 0, '=EĞER(DOĞRU;1;2)', { locale: TR })) applyOp(wb, op);
  assert.equal(getCell(wb.sheets[0], 1, 0)!.formula!.text, 'IF(TRUE,1,2)');
  wb.definedNames = [{ name: 'Doğru', value: 'Data!$B$1' }];
  for (const op of setCellInput(wb, 0, 2, 0, '=EĞER(doğru>1;1;YANLIŞ)', { locale: TR })) applyOp(wb, op);
  assert.equal(getCell(wb.sheets[0], 2, 0)!.formula!.text, 'IF(doğru>1,1,FALSE)');
});

test('formula locale: quotes doubled inside text, and error constants, come through both ways', () => {
  assert.equal(formulaFromLocal('BİRLEŞTİR("x""y;";"a,b";1,5)', TR), 'CONCATENATE("x""y;","a,b",1.5)');
  assert.equal(formulaToLocal('CONCATENATE("x""y,",1.5,"""")', TR), 'CONCATENATE("x""y,";1,5;"""")');
  assert.equal(formulaFromLocal('"a"";"&"b";1', TR), '"a"";"&"b",1');
  assert.equal(formulaFromLocal("'It''s; 1,5'!A1;2", TR), "'It''s; 1,5'!A1,2");
  for (const [stored, typed] of [
    ['IF(A1=#DIV/0!,#N/A,1.5)', 'IF(A1=#DIV/0!;#N/A;1,5)'],
    ['IFERROR(#VALUE!,#REF!)+#NAME?+#NUM!+#NULL!', 'IFERROR(#VALUE!;#REF!)+#NAME?+#NUM!+#NULL!'],
    ['SUM(#N/A,0.5,#DIV/0!,2)', 'SUM(#N/A;0,5;#DIV/0!;2)'],
  ]) {
    assert.equal(formulaToLocal(stored, TR), typed, stored);
    assert.equal(formulaFromLocal(typed, TR), stored, typed);
  }
});

test('formula locale: Turkish error constants are read when typed; what is shown stays English', () => {
  const pairs: [string, string][] = [
    ['#YOK', '#N/A'],
    ['#DEĞER!', '#VALUE!'],
    ['#BAŞV!', '#REF!'],
    ['#SAYI/0!', '#DIV/0!'],
    ['#AD?', '#NAME?'],
    ['#SAYI!', '#NUM!'],
    ['#BOŞ!', '#NULL!'],
  ];
  for (const [typed, stored] of pairs) {
    assert.deepEqual(parseCellInput(typed, { locale: TR }), { kind: 'error', value: stored }, typed);
    assert.deepEqual(parseCellInput(stored, { locale: TR }), { kind: 'error', value: stored }, stored);
    // In English the Turkish word is text.
    assert.equal(parseCellInput(typed, { locale: LOCALE_EN }).kind, 'string', typed);
    assert.equal(formulaFromLocal(`EĞERHATA(A1;${typed})`, TR), `IFERROR(A1,${stored})`, typed);
  }
  assert.deepEqual(parseCellInput('#değer!', { locale: TR }), { kind: 'error', value: '#VALUE!' });
  assert.deepEqual(parseCellInput('#sayı/0!', { locale: TR }), { kind: 'error', value: '#DIV/0!' });
  assert.deepEqual(parseCellInput(' #yok ', { locale: TR }), { kind: 'error', value: '#N/A' });
  assert.equal(parseCellInput('#YOKTUR', { locale: TR }).kind, 'string');
  assert.equal(formulaFromLocal('EĞER(A1=#SAYI/0!;#yok;1,5)', TR), 'IF(A1=#DIV/0!,#N/A,1.5)');
  assert.equal(formulaToLocal('IF(A1=#DIV/0!,#N/A,1.5)', TR), 'IF(A1=#DIV/0!;#N/A;1,5)');
  assert.equal(cellEditText({ kind: 'error', value: '#N/A', style: 0 }, 'General', { locale: TR }), '#N/A');
});

test('formula locale: Turkish function names are read, whatever their case; shown formulas keep the English names', () => {
  const cases: [string, string][] = [
    ['TOPLA(A1;A2)', 'SUM(A1,A2)'],
    ['topla(a1;a2)', 'SUM(a1,a2)'],
    ['EĞER(A1>1,5;"büyük";"küçük")', 'IF(A1>1.5,"büyük","küçük")'],
    ['eğer(ve(A1>0;değil(B1=2));doğru;YANLIŞ)', 'IF(AND(A1>0,NOT(B1=2)),TRUE,FALSE)'],
    ['DÜŞEYARA(A1;Liste!$A$1:$B$9;2;YANLIŞ)', 'VLOOKUP(A1,Liste!$A$1:$B$9,2,FALSE)'],
    ['İNDİS(A1:B9;KAÇINCI(D1;A1:A9;0);2)', 'INDEX(A1:B9,MATCH(D1,A1:A9,0),2)'],
    ['indis(A1:B9;1;1)+INDIS(A1:B9;1;2)', 'INDEX(A1:B9,1,1)+INDEX(A1:B9,1,2)'],
    ['TOPLA.ÇARPIM(A1:A3;B1:B3)', 'SUMPRODUCT(A1:A3,B1:B3)'],
    ['BAĞ_DEĞ_DOLU_SAY(A:A)-BAĞ_DEĞ_SAY(A:A)', 'COUNTA(A:A)-COUNT(A:A)'],
    ['YUVARLA(ORTALAMA(A1:A9);1)&" "&METNEÇEVİR(BUGÜN();"gg.aa.yyyy")', 'ROUND(AVERAGE(A1:A9),1)&" "&TEXT(TODAY(),"gg.aa.yyyy")'],
    ['MAK(A1;MİN(B1;C1))+mod(7;2)+MUTLAK(-1)+TAMSAYI(1,9)', 'MAX(A1,MIN(B1,C1))+MOD(7,2)+ABS(-1)+INT(1.9)'],
    ['EĞERHATA(ETOPLA(A:A;">0";B:B)/EĞERSAY(A:A;">0");0)', 'IFERROR(SUMIF(A:A,">0",B:B)/COUNTIF(A:A,">0"),0)'],
    ['ÇOKETOPLA(C:C;A:A;"x";B:B;1)', 'SUMIFS(C:C,A:A,"x",B:B,1)'],
    ['SOLDAN(A1;2)&SAĞDAN(A1;2)&PARÇAAL(A1;2;3)&UZUNLUK(A1)&BİRLEŞTİR(A1;B1)', 'LEFT(A1,2)&RIGHT(A1,2)&MID(A1,2,3)&LEN(A1)&CONCATENATE(A1,B1)'],
    ['TARİH(YIL(A1);AY(A1);GÜN(A1))+ŞİMDİ()', 'DATE(YEAR(A1),MONTH(A1),DAY(A1))+NOW()'],
    ['BOŞLUKSAY(A1:A9)+YADA(A1;B1)', 'COUNTBLANK(A1:A9)+OR(A1,B1)'],
  ];
  for (const [typed, stored] of cases) {
    assert.equal(formulaFromLocal(typed, TR), stored, typed);
    // What is stored can be read by the engine's parser.
    assert.doesNotThrow(() => parseFormula(stored), stored);
  }
  // A name is only a function where it is called: a defined name or a sheet called AY stays.
  assert.equal(formulaFromLocal('AY*2+AY!A1+TOPLA', TR), 'AY*2+AY!A1+TOPLA');
  assert.equal(formulaFromLocal('"TOPLA(1;2)"&\'EĞER(\'!A1', TR), '"TOPLA(1;2)"&\'EĞER(\'!A1');
  // English names stay, in English as in Turkish; other locales do not read Turkish names.
  assert.equal(formulaFromLocal('SUM(A1;A2)+TOPLA(B1)', TR), 'SUM(A1,A2)+SUM(B1)');
  assert.equal(formulaFromLocal('TOPLA(A1,A2)', LOCALE_EN), 'TOPLA(A1,A2)');
  assert.equal(formulaToLocal('SUM(A1,A2)', TR), 'SUM(A1;A2)');
});

test('formula locale: a locale with a decimal comma that is not Turkish translates separators only', () => {
  const german = { ...LOCALE_EN, decimal: ',', group: '.' };
  assert.equal(formulaFromLocal('TOPLA(1,5;2)', german), 'TOPLA(1.5,2)');
  assert.equal(formulaToLocal('SUM(1.5,2)', german), 'SUM(1,5;2)');
});

test('formula locale: very long and odd text comes back as it is, quickly', () => {
  const long = '1,'.repeat(40000);
  assert.equal(formulaFromLocal(long, TR), long);
  const big = `SUM(${'A1;'.repeat(8000)}1)`;
  const started = Date.now();
  assert.equal(formulaFromLocal(big, TR).length, big.length);
  assert.ok(Date.now() - started < 1000);
  for (const odd of ['', '(', '"', "'", '[', '{', '}}}', '#', ';;', '\\', '{\\', '$', '1E', '1,5E+', '"a,', "'a,", '[a,']) {
    assert.equal(typeof formulaFromLocal(odd, TR), 'string');
    assert.equal(typeof formulaToLocal(odd, TR), 'string');
  }
  // Nothing but a RangeError ever comes out of text that makes no sense.
  for (const odd of ['1,', ',1', ',', ',,,', '{,}', '#,', 'DOĞRU,']) {
    assert.throws(() => formulaFromLocal(odd, TR), RangeError);
    assert.equal(typeof formulaToLocal(odd, TR), 'string');
    assert.equal(parseCellInput(`=${odd}`, { locale: TR }).invalid, true);
  }
  assert.equal(formulaFromLocal(undefined as unknown as string, TR), undefined);
});

test('formula locale: typed input is stored in the file notation and shown again in the locale', () => {
  assert.deepEqual(parseCellInput('=TOPLA(A1;B1)*1,5', { locale: TR }), { kind: 'blank', value: null, formula: 'SUM(A1,B1)*1.5' });
  assert.deepEqual(parseCellInput('=SUM(A1,B1)*1.5', { locale: LOCALE_EN }), { kind: 'blank', value: null, formula: 'SUM(A1,B1)*1.5' });
  assert.deepEqual(parseCellInput('=SUM(A1;B1)', { locale: LOCALE_EN }), { kind: 'blank', value: null, formula: 'SUM(A1;B1)' });
  const cell: Cell = { kind: 'number', value: 3, style: 0, formula: { text: 'SUM(A1,B1)*1.5', type: 'normal' } };
  assert.equal(cellEditText(cell, 'General', { locale: TR }), '=SUM(A1;B1)*1,5');
  assert.equal(cellEditText(cell, 'General', { locale: LOCALE_EN }), '=SUM(A1,B1)*1.5');
  assert.equal(cellEditText(cell, 'General'), '=SUM(A1,B1)*1.5');
});

test('formula locale: the commands store, show, refuse and paste formulas in the locale', () => {
  const wb = workbook();
  const sheet = wb.sheets[0];
  for (const op of setCellInput(wb, 0, 0, 0, '=EĞER(B1>1,5;TOPLA(B1;C1);0)', { locale: TR })) applyOp(wb, op);
  assert.equal(getCell(sheet, 0, 0)!.formula!.text, 'IF(B1>1.5,SUM(B1,C1),0)');
  assert.equal(editTextAt(wb, sheet, 0, 0, { locale: TR }), '=IF(B1>1,5;SUM(B1;C1);0)');
  assert.equal(editTextAt(wb, sheet, 0, 0, { locale: LOCALE_EN }), '=IF(B1>1.5,SUM(B1,C1),0)');
  // Typing back what the editor shows changes nothing.
  for (const op of setCellInput(wb, 0, 0, 0, editTextAt(wb, sheet, 0, 0, { locale: TR }), { locale: TR })) applyOp(wb, op);
  assert.equal(getCell(sheet, 0, 0)!.formula!.text, 'IF(B1>1.5,SUM(B1,C1),0)');

  // What cannot be read after the translation is refused, as before.
  assert.throws(() => setCellInput(wb, 0, 1, 0, '=TOPLA(A1;;', { locale: TR }), /The formula cannot be read/);
  assert.throws(() => setCellInput(wb, 0, 1, 0, '=SUM(A1;B1)', { locale: LOCALE_EN }), /The formula cannot be read/);

  // Pasted from elsewhere: translated like typed text. Pasted from this workbook: the cells themselves.
  let result = pasteAt(wb, 0, { top: 2, left: 0, bottom: 2, right: 0 }, '=TOPLA(B3;C3)\t=1,5*2\t1,5', undefined, { locale: TR });
  for (const op of result.ops) applyOp(wb, op);
  assert.equal(getCell(sheet, 2, 0)!.formula!.text, 'SUM(B3,C3)');
  assert.equal(getCell(sheet, 2, 1)!.formula!.text, '1.5*2');
  assert.equal(getCell(sheet, 2, 2)!.value, 1.5);
  const copied = copyRange(wb, 0, { top: 2, left: 0, bottom: 2, right: 1 }, { locale: TR });
  result = pasteAt(wb, 0, { top: 4, left: 0, bottom: 4, right: 0 }, copied.text, copied, { locale: TR });
  for (const op of result.ops) applyOp(wb, op);
  assert.equal(getCell(sheet, 4, 0)!.formula!.text, 'SUM(B5,C5)');
  assert.equal(getCell(sheet, 4, 1)!.formula!.text, '1.5*2');
});
