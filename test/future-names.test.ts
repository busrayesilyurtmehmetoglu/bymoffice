// The prefixes files give to newer functions, to the names of LET and LAMBDA, and the calls
// they store for "A1#" and "@".

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { parseFormula } from '../src/calc/parse';
import { withFuturePrefixes, withoutFuturePrefixes } from '../src/edit/future-names';
import { formulaFromLocal, formulaToLocal } from '../src/edit/formula-locale';
import { LOCALE_EN, LOCALE_TR } from '../src/format/locale';
import { open } from '../src/index';
import { replaceAllOps } from '../src/ops/find';
import { buildXlsx, worksheetXml } from './fixtures';

const CASES: [shown: string, stored: string][] = [
  ['SUM(A1:A3)', 'SUM(A1:A3)'],
  ['XLOOKUP(A1,B:B,C:C)', '_xlfn.XLOOKUP(A1,B:B,C:C)'],
  ['FILTER(A1:A9,B1:B9>1)', '_xlfn._xlws.FILTER(A1:A9,B1:B9>1)'],
  ['SORT(UNIQUE(A1:A9))', '_xlfn._xlws.SORT(_xlfn.UNIQUE(A1:A9))'],
  ['SORTBY(A1:A9,B1:B9,-1)', '_xlfn.SORTBY(A1:A9,B1:B9,-1)'],
  ['SEQUENCE(3)*2', '_xlfn.SEQUENCE(3)*2'],
  ['LET(x,1,y,x+1,x*y)', '_xlfn.LET(_xlpm.x,1,_xlpm.y,_xlpm.x+1,_xlpm.x*_xlpm.y)'],
  ['LAMBDA(x,x+1)(2)', '_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)(2)'],
  ['LET(f,LAMBDA(n,n*n),f(3))', '_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.n,_xlpm.n*_xlpm.n),_xlpm.f(3))'],
  ['MAP(A1:A3,LAMBDA(v,v*2))', '_xlfn.MAP(A1:A3,_xlfn.LAMBDA(_xlpm.v,_xlpm.v*2))'],
  ['A1#', '_xlfn.ANCHORARRAY(A1)'],
  ['SUM(Sheet2!B3#)', 'SUM(_xlfn.ANCHORARRAY(Sheet2!B3))'],
  ["COUNT('My Sheet'!$C$1#)", "COUNT(_xlfn.ANCHORARRAY('My Sheet'!$C$1))"],
  ['A1#*10', '_xlfn.ANCHORARRAY(A1)*10'],
  ['@A1:A3*2', '_xlfn.SINGLE(A1:A3)*2'],
  ['SUM(@A1:A3)', 'SUM(_xlfn.SINGLE(A1:A3))'],
  ['@A1#', '_xlfn.SINGLE(_xlfn.ANCHORARRAY(A1))'],
  // A name of LET is a parameter only where it is bound; a defined name of the same spelling elsewhere stays.
  ['x+LET(x,2,x)', 'x+_xlfn.LET(_xlpm.x,2,_xlpm.x)'],
];

test('a formula as shown and as stored', () => {
  for (const [shown, stored] of CASES) {
    assert.equal(withFuturePrefixes(shown), stored, `stored: ${shown}`);
    assert.equal(withoutFuturePrefixes(stored), shown, `shown: ${stored}`);
    // Both read as the same formula.
    assert.deepEqual(parseFormula(stored), parseFormula(shown), `parsed: ${shown}`);
  }
});

test('what has its prefix already, and what needs none, stays as it is', () => {
  for (const [, stored] of CASES) assert.equal(withFuturePrefixes(stored), stored);
  for (const text of ['IF(A1="LET(x,1,x)",1,2)', 'SUM(1E+5,2)', "'A#'!B1", '#N/A', 'IFERROR(A1,#REF!)', '{1,2;3,4}', 'Table1[#Headers]', 'Sheet1:Sheet3!A1', '1:3', '"@"&A1']) {
    assert.equal(withFuturePrefixes(text), text, text);
    assert.equal(withoutFuturePrefixes(text), text, text);
  }
});

test('long formulas are translated in about the time it takes to read them', () => {
  const many = Array.from({ length: 700 }, (_, k) => `A${k + 1}#`).join('+');
  const started = Date.now();
  const stored = withFuturePrefixes(many);
  assert.equal(withoutFuturePrefixes(stored), many);
  const words = `LET(x,1,${Array.from({ length: 2000 }, () => 'x').join('+')})`;
  assert.equal(withoutFuturePrefixes(withFuturePrefixes(words)), words);
  // An unclosed call is shown as it came, without its prefix.
  assert.equal(withoutFuturePrefixes('SUM(_xlfn.ANCHORARRAY(A1'), 'SUM(ANCHORARRAY(A1');
  assert.ok(Date.now() - started < 2000, `${Date.now() - started} ms`);
});

test('typed in a locale: Turkish names and separators, then the prefixes', () => {
  assert.equal(formulaFromLocal('FİLTRE(A1:A9;B1:B9>1,5)', LOCALE_TR), '_xlfn._xlws.FILTER(A1:A9,B1:B9>1.5)');
  assert.equal(formulaFromLocal('LET(x;1;x+1)', LOCALE_TR), '_xlfn.LET(_xlpm.x,1,_xlpm.x+1)');
  assert.equal(formulaFromLocal('BENZERSİZ(A1#)', LOCALE_TR), '_xlfn.UNIQUE(_xlfn.ANCHORARRAY(A1))');
  assert.equal(formulaToLocal('_xlfn.LET(_xlpm.x,1.5,_xlpm.x+1)', LOCALE_TR), 'LET(x;1,5;x+1)');
  assert.equal(formulaFromLocal('SEQUENCE(3)', LOCALE_EN), '_xlfn.SEQUENCE(3)');
  assert.equal(formulaToLocal('_xlfn.XLOOKUP(A1,B:B,C:C)', LOCALE_EN), 'XLOOKUP(A1,B:B,C:C)');
});

test('a prefix this engine does not know the function of is shown, and so kept', () => {
  for (const stored of ['_xlfn.CHISQ.DIST(A1,2,TRUE)+_xlfn.REGEXTEST("a",",")', '_xlfn.GROUPBY(A1:A2,B1:B2,C1)', '_xlfn.FILTER(A1:A3,B1:B3)', '_XLFN.XLOOKUP(1,A:A,B:B)', "_xlfn.abc!A1+'_xlfn.x'!B2", '_xlpm.x!A1']) {
    const shown = withoutFuturePrefixes(stored);
    assert.equal(withFuturePrefixes(shown), stored, `${stored} -> ${shown}`);
    for (const locale of [LOCALE_EN, LOCALE_TR]) assert.equal(formulaFromLocal(formulaToLocal(stored, locale), locale), stored, `${stored} in ${locale.decimal}`);
  }
  // A sheet whose name starts like a prefix keeps its name.
  assert.equal(withoutFuturePrefixes('_xlfn.abc!A1'), '_xlfn.abc!A1');
});

test('"@" and "#" are shown short only where that reads as the same formula, and every round trip is the same', () => {
  const cases: [stored: string, shown: string][] = [
    ['_xlfn.SINGLE(INDEX(A:A,1))', '@INDEX(A:A,1)'],
    ['_xlfn.SINGLE(-A1)', '@(-A1)'],
    ['_xlfn.SINGLE("x")', '@"x"'],
    ['_xlfn.SINGLE({1,2})', '@{1,2}'],
    ['_xlfn.SINGLE(A1:A3) B2:B4', '@A1:A3 B2:B4'],
    ['_xlfn.SINGLE(A1+1)', '@(A1+1)'],
    // "#" is read after a reference or a name only.
    ['_xlfn.ANCHORARRAY(OFFSET(A1,1,1))', 'ANCHORARRAY(OFFSET(A1,1,1))'],
    ['_xlfn.ANCHORARRAY(A1+1)', 'ANCHORARRAY(A1+1)'],
    ['SUM(_xlfn.SINGLE(_xlfn.ANCHORARRAY(A1)))', 'SUM(@A1#)'],
  ];
  for (const [stored, shown] of cases) {
    assert.equal(withoutFuturePrefixes(stored), shown, stored);
    let text = stored;
    for (let k = 0; k < 3; k++) text = withFuturePrefixes(withoutFuturePrefixes(text));
    assert.equal(text, stored, `round trips of ${stored}`);
    assert.deepEqual(parseFormula(shown), parseFormula(stored), `${shown} reads as ${stored}`);
  }
  // The spill of a corner of a range of another sheet is of that sheet.
  assert.deepEqual(parseFormula('Sheet1!A1:B2#'), parseFormula('Sheet1!A1:_xlfn.ANCHORARRAY(Sheet1!B2)'));
});

test('replacing in formulas keeps the prefixes the shown text does not have', async () => {
  const bytes = await buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>_xlfn.CHISQ.DIST(A1,2,TRUE)+_xlfn.XLOOKUP(2,A1:A2,A1:A2)</f><v>0.3</v></c></row></sheetData>') }] });
  const workbook = await (await open(bytes)).workbook();
  const result = replaceAllOps(workbook, ',2,', ',3,', {});
  const op = result.ops[0] as { cells: { cell: { formula: { text: string } } }[] };
  assert.equal(op.cells[0].cell.formula.text, '_xlfn.CHISQ.DIST(A1,3,TRUE)+_xlfn.XLOOKUP(2,A1:A2,A1:A2)');
});

test('replacing in many cells of a shared formula works the group out once', async () => {
  const count = 20000;
  let rows = '';
  for (let r = 1; r <= count; r++) rows += r === 1 ? `<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B${count}" si="0">A1+1</f><v>2</v></c></row>` : `<row r="${r}"><c r="A${r}"><v>1</v></c><c r="B${r}"><f t="shared" si="0"/><v>2</v></c></row>`;
  const workbook = await (await open(await buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml(`<sheetData>${rows}</sheetData>`) }] }))).workbook();
  const started = Date.now();
  const result = replaceAllOps(workbook, '+1', '+2', {});
  assert.equal(result.replaced, count);
  // A guard against work for each cell over the whole sheet (13 seconds before), not a measure of speed.
  assert.ok(Date.now() - started < 5000, `${Date.now() - started} ms`);
});

test('what is shown is typed back as the very text stored: the whole formula decides, parentheses included', () => {
  const cases: [stored: string, shown: string][] = [
    // "@A1:B5" would be the intersection of the whole range: the call stays a call.
    ['_xlfn.SINGLE(A1):B5', '@(A1):B5'],
    ['SUM(_xlfn.SINGLE(A1):B5)', 'SUM(@(A1):B5)'],
    ['_xlfn.SINGLE(_xlfn.SINGLE(A1):B5)', '@(@(A1):B5)'],
    ['SUM(_xlfn.SINGLE(_xlfn.ANCHORARRAY(A1)):B5)', 'SUM(@(ANCHORARRAY(A1)):B5)'],
    ['_xlfn.SINGLE(A1:B5)', '@A1:B5'],
    // Parentheses of the file are kept.
    ['_xlfn.SINGLE(((A1)))', '@(((A1)))'],
    ['_xlfn.SINGLE((-(A1)))', '@((-(A1)))'],
    ['-_xlfn.SINGLE(({1,2})):_xlfn.SEQUENCE((TRUE))', '-@(({1,2})):SEQUENCE((TRUE))'],
  ];
  for (const [stored, shown] of cases) {
    assert.equal(withoutFuturePrefixes(stored), shown, stored);
    assert.equal(withFuturePrefixes(withoutFuturePrefixes(stored)), stored, stored);
  }
});

test('a stored formula that the shown text would not give back is shown as stored', () => {
  // A name of LET outside any LET: without its prefix it would be typed back as another name.
  const stored = 'SUM(_xlpm.x)';
  assert.equal(withFuturePrefixes(withoutFuturePrefixes(stored)), stored);
  // Every text of a small grammar comes back as it was.
  const parts = ['A1', 'B2:C3', '(A1)', '((B2))', '1', '"t"', '{1,2}', 'A1#', '@A1', 'SUM(A1)', '-A1'];
  const wrap = ['_xlfn.SINGLE(%)', '_xlfn.ANCHORARRAY(%)', '%:B5', 'SUM(%)', '(%)', '-%'];
  let tried = 0;
  for (const a of parts) {
    for (const w1 of wrap) {
      for (const w2 of wrap) {
        const text = withFuturePrefixes(w2.replace('%', w1.replace('%', a)));
        try {
          parseFormula(text);
        } catch (error) {
          continue;
        }
        tried++;
        assert.equal(withFuturePrefixes(withoutFuturePrefixes(text)), text, text);
      }
    }
  }
  assert.ok(tried > 200, `${tried}`);
});
