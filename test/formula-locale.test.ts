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

test('formula locale: a formula typed the English way is understood, except where a comma stands between digits', () => {
  assert.equal(formulaFromLocal('SUM(A1,B1)', TR), 'SUM(A1,B1)');
  assert.equal(formulaFromLocal('ROUND(A1,2)', TR), 'ROUND(A1,2)');
  assert.equal(formulaFromLocal('A1*1.5', TR), 'A1*1.5');
  assert.equal(formulaFromLocal('IF(A1>1, 2, 3)', TR), 'IF(A1>1, 2, 3)');
  // As in the Turkish Excel: this is the one number 1,5.
  assert.equal(formulaFromLocal('SUM(1,5)', TR), 'SUM(1.5)');
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
  for (const odd of ['', '(', '"', "'", '[', '{', '}}}', '#', '1,', ',1', ';;', '\\', '{\\', '$', '1E', '1,5E+']) {
    assert.equal(typeof formulaFromLocal(odd, TR), 'string');
    assert.equal(typeof formulaToLocal(odd, TR), 'string');
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
