// Tests for what completes the function library: the rest of the trigonometric family,
// further means and deviations, functions of two lists, the Turkish case of letters, and
// every family typed the Turkish way (Turkish names, ";" between arguments, "," in numbers).
//
// Expected numbers were worked out apart from the engine, with the standard library of
// Python (math, statistics, datetime), or are the examples of Excel's own documentation
// where a comment says so.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine, type CalcEngineOptions } from '../src/calc/engine';
import { isBuiltinForm } from '../src/calc/eval';
import { FUNCTIONS } from '../src/calc/functions/index';
import { formulaFromLocal, formulaToLocal } from '../src/edit/formula-locale';
import { LOCALE_EN, LOCALE_TR, type FormatLocale } from '../src/format/locale';
import { setCellInput } from '../src/ops/commands';
import { viewEngineOptions } from '../src/ui/engine-options';
import { createSheet, emptyStyleSheet, getCell, parseCellRef, type Cell, type CellFormula, type Row, type Workbook } from '../src/model/sheet/types';

type Plain = number | string | boolean | null;
type Cells = Record<string, Plain>;

function makeCell(spec: Plain): Cell {
  if (typeof spec === 'string' && spec.charAt(0) === '=') {
    const formula: CellFormula = { text: spec.slice(1), type: 'normal' };
    return { kind: 'blank', value: null, style: 0, formula };
  }
  return { kind: spec === null ? 'blank' : typeof spec === 'number' ? 'number' : typeof spec === 'boolean' ? 'boolean' : 'string', value: spec, style: 0 };
}

function book(cells: Cells = {}, date1904 = false): Workbook {
  const workbook: Workbook = { sheets: [], styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904, activeSheet: 0 };
  const sheet = createSheet({ name: 'S', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  for (const address of Object.keys(cells)) {
    const at = parseCellRef(address) as { row: number; col: number };
    let row = sheet.rows.get(at.row);
    if (!row) {
      row = { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] } as Row;
      sheet.rows.set(at.row, row);
    }
    row.cells[at.col] = makeCell(cells[address]);
    if (at.row >= sheet.rowCount) sheet.rowCount = at.row + 1;
    if (at.col >= sheet.colCount) sheet.colCount = at.col + 1;
  }
  workbook.sheets.push(sheet);
  return workbook;
}

function formulaToLocalText(stored: string): string {
  return formulaToLocal(stored, LOCALE_TR);
}

/** The formula of the cell a list of operations puts down. */
function cellOf(ops: unknown): string | undefined {
  const found = /"text":"((?:[^"\\]|\\.)*)"/.exec(JSON.stringify(ops));
  return found ? (JSON.parse('"' + found[1] + '"') as string) : undefined;
}

const NO_LIMIT: CalcEngineOptions = { budget: { milliseconds: 0 } };
/** A cell far from the data, for formulas that are only asked. */
const AWAY = { sheet: 0, row: 50, col: 20 };

function engineOf(cells: Cells = {}, options: CalcEngineOptions = NO_LIMIT): CalcEngine {
  return new CalcEngine(book(cells), options);
}

const blank = engineOf();

function calc(formula: string, engine: CalcEngine = blank): unknown {
  return engine.evaluateText(formula, AWAY);
}

function near(actual: unknown, expected: number, formula: string): void {
  assert.equal(typeof actual, 'number', `${formula} gives ${JSON.stringify(actual)}, not a number`);
  const tolerance = Math.max(1e-12, Math.abs(expected) * 1e-10);
  assert.ok(Math.abs((actual as number) - expected) <= tolerance, `${formula} gives ${actual}, expected ${expected}`);
}

function isErrorOf(actual: unknown, code: string, formula: string): void {
  assert.deepEqual(actual, { error: code }, formula);
}

// ---------------------------------------------------------------------------
// Trigonometry

test('the rest of the trigonometric family', () => {
  const cases: [string, number][] = [
    ['ASINH(1)', 0.881373587019543],
    ['ASINH(0)', 0],
    ['ASINH(-1)', -0.881373587019543],
    ['ACOSH(2)', 1.3169578969248166],
    ['ACOSH(1)', 0],
    ['ATANH(0.5)', 0.5493061443340549],
    ['ATANH(0)', 0],
    ['COT(1)', 0.6420926159343308],
    ['CSC(1)', 1.1883951057781212],
    ['SEC(1)', 1.8508157176809255],
    ['SEC(0)', 1],
    ['COTH(1)', 1.3130352854993315],
    ['CSCH(1)', 0.8509181282393216],
    ['SECH(1)', 0.6480542736638855],
    ['SECH(0)', 1],
    // Excel's documentation: ACOT(2) is 0.4636.
    ['ACOT(2)', 0.4636476090008059],
    ['ACOT(0)', Math.PI / 2],
    ['ACOT(-1)', 2.356194490192345],
    ['ACOTH(2)', 0.5493061443340549],
    ['ACOTH(-3)', -0.34657359027997264],
    // Excel's documentation: SQRTPI(2) is 2.506628.
    ['SQRTPI(2)', 2.5066282746310002],
    ['SQRTPI(0)', 0],
    // Text that reads as a number and booleans are converted, as by the other functions of one number.
    ['COT("1")', 0.6420926159343308],
    ['SEC(TRUE)', 1.8508157176809255],
  ];
  for (const [formula, expected] of cases) near(calc(formula), expected, formula);
});

test('trigonometric functions outside their domain', () => {
  const cases: [string, string][] = [
    ['ACOSH(0.5)', '#NUM!'],
    ['ATANH(1)', '#NUM!'],
    ['ATANH(-1)', '#NUM!'],
    ['ATANH(2)', '#NUM!'],
    ['ACOTH(1)', '#NUM!'],
    ['ACOTH(0)', '#NUM!'],
    ['ACOTH(-1)', '#NUM!'],
    ['COT(0)', '#DIV/0!'],
    ['CSC(0)', '#DIV/0!'],
    ['COTH(0)', '#DIV/0!'],
    ['CSCH(0)', '#DIV/0!'],
    // 2^27 and beyond is not an angle to Excel.
    ['COT(134217728)', '#NUM!'],
    ['CSC(-134217728)', '#NUM!'],
    ['SEC(1E+10)', '#NUM!'],
    ['SQRTPI(-1)', '#NUM!'],
    ['SEC("abc")', '#VALUE!'],
    ['ASINH("abc")', '#VALUE!'],
    ['COT(#N/A)', '#N/A'],
    ['ACOT(1/0)', '#DIV/0!'],
  ];
  for (const [formula, code] of cases) isErrorOf(calc(formula), code, formula);
  // Just inside the limit an answer is given.
  assert.equal(typeof calc('COT(134217727)'), 'number', 'COT just below 2^27');
});

test('trigonometric functions read cells: an empty cell is 0, a range gives the cell of the row', () => {
  const engine = engineOf({ A1: 1, A2: 'x', A4: '=1/0' });
  near(calc('ASINH(A1)', engine), 0.881373587019543, 'ASINH(A1)');
  assert.equal(calc('ASINH(A3)', engine), 0, 'an empty cell is 0');
  near(calc('SEC(A3)', engine), 1, 'SEC of an empty cell');
  isErrorOf(calc('COT(A3)', engine), '#DIV/0!', 'COT of an empty cell');
  isErrorOf(calc('SECH(A2)', engine), '#VALUE!', 'SECH of text');
  engine.recalculate();
  isErrorOf(calc('SECH(A4)', engine), '#DIV/0!', 'the error of a cell is the result');
  // An array argument gives an array; SUM adds its elements.
  near(calc('SUM(SECH({0,0}))', engine), 2, 'SUM(SECH({0,0}))');
});

test('FACTDOUBLE', () => {
  // Excel's documentation: FACTDOUBLE(6) is 48, FACTDOUBLE(7) is 105.
  assert.equal(calc('FACTDOUBLE(6)'), 48);
  assert.equal(calc('FACTDOUBLE(7)'), 105);
  assert.equal(calc('FACTDOUBLE(0)'), 1);
  assert.equal(calc('FACTDOUBLE(1)'), 1);
  assert.equal(calc('FACTDOUBLE(-1)'), 1);
  assert.equal(calc('FACTDOUBLE(6.9)'), 48, 'the number is cut to an integer');
  near(calc('FACTDOUBLE(300)'), 8.154414069380594e307, 'FACTDOUBLE(300)');
  isErrorOf(calc('FACTDOUBLE(301)'), '#NUM!', 'too large for a number');
  isErrorOf(calc('FACTDOUBLE(-2)'), '#NUM!', 'FACTDOUBLE(-2)');
  isErrorOf(calc('FACTDOUBLE("a")'), '#VALUE!', 'FACTDOUBLE("a")');
  isErrorOf(calc('FACTDOUBLE(#REF!)'), '#REF!', 'FACTDOUBLE(#REF!)');
});

// ---------------------------------------------------------------------------
// Means and deviations

test('HARMEAN, AVEDEV and DEVSQ', () => {
  // Excel's documentation: HARMEAN(4,5,8,7,11,4,3) is 5.028376, DEVSQ of the same numbers is 48.
  near(calc('HARMEAN(4,5,8,7,11,4,3)'), 5.028375962061728, 'HARMEAN');
  near(calc('DEVSQ(4,5,8,7,11,4,3)'), 48, 'DEVSQ');
  near(calc('AVEDEV(4,5,8,7,11,4,3)'), 2.2857142857142856, 'AVEDEV');
  // Excel's documentation: AVEDEV(4,5,6,7,5,4,3) is 1.020408.
  near(calc('AVEDEV(4,5,6,7,5,4,3)'), 1.0204081632653061, 'AVEDEV of the documented example');
  near(calc('HARMEAN({4,5;8,7},11,4,3)'), 5.028375962061728, 'HARMEAN of an array and numbers');
  assert.equal(calc('HARMEAN(5)'), 5);
  assert.equal(calc('AVEDEV(5)'), 0);
  assert.equal(calc('DEVSQ(5)'), 0);
  isErrorOf(calc('HARMEAN(4,0,5)'), '#NUM!', 'a zero');
  isErrorOf(calc('HARMEAN(4,-1)'), '#NUM!', 'a negative number');
  isErrorOf(calc('HARMEAN("abc")'), '#VALUE!', 'text given directly');
  isErrorOf(calc('AVEDEV(1,#N/A)'), '#N/A', 'an error given directly');
});

test('HARMEAN, AVEDEV and DEVSQ over cells: text, booleans and empty cells are passed over', () => {
  const engine = engineOf({ A1: 4, A2: 5, A3: 'text', A4: true, A6: 8, A7: 7, A8: 11, A9: 4, A10: 3, B1: 'only', B2: 'text', C1: 1, C2: '=1/0' });
  engine.recalculate();
  near(calc('HARMEAN(A1:A10)', engine), 5.028375962061728, 'HARMEAN(A1:A10)');
  near(calc('AVEDEV(A1:A10)', engine), 2.2857142857142856, 'AVEDEV(A1:A10)');
  near(calc('DEVSQ(A:A)', engine), 48, 'DEVSQ(A:A)');
  // UNSURE: without any number these give #NUM! here; Excel was not asked.
  isErrorOf(calc('HARMEAN(B1:B2)', engine), '#NUM!', 'HARMEAN without numbers');
  isErrorOf(calc('AVEDEV(B1:B2)', engine), '#NUM!', 'AVEDEV without numbers');
  isErrorOf(calc('DEVSQ(E1:E5)', engine), '#NUM!', 'DEVSQ of empty cells');
  isErrorOf(calc('HARMEAN(C1:C2)', engine), '#DIV/0!', 'an error inside the range');
});

test('STDEVA, STDEVPA, VARA and VARPA count text as 0 and booleans as 1 and 0', () => {
  const engine = engineOf({ A1: 1, A2: 2, A3: 'abc', A4: true, A5: 5, B1: 7 });
  // Python: statistics over [1, 2, 0, 1, 5].
  near(calc('STDEVA(A1:A5)', engine), 1.9235384061671346, 'STDEVA');
  near(calc('STDEVPA(A1:A5)', engine), 1.7204650534085253, 'STDEVPA');
  near(calc('VARA(A1:A5)', engine), 3.7, 'VARA');
  near(calc('VARPA(A1:A5)', engine), 2.96, 'VARPA');
  // The plain functions leave the text and the boolean out: [1, 2, 5].
  near(calc('VAR(A1:A5)', engine), 4.333333333333333, 'VAR');
  // An empty cell is left out by both.
  near(calc('VARPA(A1:A7)', engine), 2.96, 'VARPA with empty cells');
  isErrorOf(calc('STDEVA(B1)', engine), '#DIV/0!', 'one value has no sample deviation');
  assert.equal(calc('STDEVPA(B1)', engine), 0, 'one value has a population deviation of 0');
  isErrorOf(calc('VARA("abc",1)', engine), '#VALUE!', 'text given directly is not a number');
  near(calc('VARA(TRUE,3)', engine), 2, 'a boolean given directly is 1');
  isErrorOf(calc('VARPA(1,#NUM!)', engine), '#NUM!', 'an error given directly');
});

// ---------------------------------------------------------------------------
// Two lists

test('SUMX2MY2, SUMX2PY2 and SUMXMY2', () => {
  // Excel's documentation: -55, 521 and 79 for these two lists.
  const x = '{2,3,9,1,8,7,5}';
  const y = '{6,5,11,7,5,4,4}';
  assert.equal(calc(`SUMX2MY2(${x},${y})`), -55);
  assert.equal(calc(`SUMX2PY2(${x},${y})`), 521);
  assert.equal(calc(`SUMXMY2(${x},${y})`), 79);
  assert.equal(calc('SUMXMY2(5,2)'), 9, 'single numbers are lists of one');
  isErrorOf(calc('SUMX2MY2({1,2,3},{1,2})'), '#N/A', 'lists of different sizes');
  isErrorOf(calc('SUMX2PY2({1,#DIV/0!},{1,2})'), '#DIV/0!', 'an error in a list');
});

test('functions of two lists take the positions where both hold a number', () => {
  const engine = engineOf({ A1: 2, A2: 3, A3: 'x', A4: 1, A5: 8, B1: 6, B2: 5, B3: 11, B5: 5, B4: true });
  // Pairs left: (2,6), (3,5), (8,5).
  assert.equal(calc('SUMX2MY2(A1:A5,B1:B5)', engine), 4 - 36 + 9 - 25 + 64 - 25);
  assert.equal(calc('SUMXMY2(A1:A5,B1:B5)', engine), 16 + 4 + 9);
  // UNSURE: no pair of numbers at all gives 0 here; Excel was not asked.
  assert.equal(calc('SUMX2PY2(A3:A4,B3:B4)', engine), 0);
  isErrorOf(calc('SUMXMY2(A1:A5,B1:B4)', engine), '#N/A', 'ranges of different sizes');
});

test('PEARSON, RSQ, COVAR and COVARIANCE', () => {
  const y = '{2,3,9,1,8,7,5}';
  const x = '{6,5,11,7,5,4,4}';
  // Excel's documentation: RSQ of these lists is 0.05795.
  near(calc(`RSQ(${y},${x})`), 0.05795019157088122, 'RSQ');
  near(calc(`PEARSON(${y},${x})`), 0.24072846024282468, 'PEARSON');
  assert.deepEqual(calc(`PEARSON(${y},${x})`), calc(`CORREL(${y},${x})`), 'PEARSON is CORREL');
  near(calc(`COVAR(${y},${x})`), 1.5714285714285714, 'COVAR');
  near(calc(`COVARIANCE.P(${y},${x})`), 1.5714285714285714, 'COVARIANCE.P');
  near(calc(`COVARIANCE.S(${y},${x})`), 1.8333333333333333, 'COVARIANCE.S');
  isErrorOf(calc('RSQ({1,2,3},{4,4,4})'), '#DIV/0!', 'a list that does not vary');
  isErrorOf(calc('RSQ({1,2,3},{4,5})'), '#N/A', 'lists of different sizes');
  isErrorOf(calc('COVARIANCE.S(1,2)'), '#DIV/0!', 'one pair has no sample covariance');
  assert.equal(calc('COVARIANCE.P(1,2)'), 0, 'one pair has a population covariance of 0');
  isErrorOf(calc('COVAR({"a","b"},{1,2})'), '#DIV/0!', 'no pair of numbers');
  isErrorOf(calc('COVAR({1,#N/A},{1,2})'), '#N/A', 'an error in a list');
  isErrorOf(calc('RSQ({1,2},{#REF!,2})'), '#REF!', 'an error in the second list');
});

test('every new function refuses a wrong number of arguments', () => {
  const one = ['ASINH', 'ACOSH', 'ATANH', 'COT', 'CSC', 'SEC', 'COTH', 'CSCH', 'SECH', 'ACOT', 'ACOTH', 'SQRTPI', 'FACTDOUBLE'];
  const two = ['SUMX2MY2', 'SUMX2PY2', 'SUMXMY2', 'PEARSON', 'RSQ', 'COVAR', 'COVARIANCE.P', 'COVARIANCE.S'];
  const context = {} as never;
  for (const name of one) {
    const spec = FUNCTIONS.get(name);
    assert.ok(spec, `${name} is registered`);
    assert.deepEqual([spec.min, spec.max], [1, 1], name);
    assert.deepEqual(spec.call([], context), { error: '#VALUE!' }, `${name} without arguments`);
    assert.deepEqual(spec.call([1, 2], context), { error: '#VALUE!' }, `${name} with two arguments`);
  }
  for (const name of two) {
    const spec = FUNCTIONS.get(name);
    assert.ok(spec, `${name} is registered`);
    assert.deepEqual([spec.min, spec.max], [2, 2], name);
    assert.deepEqual(spec.call([1], context), { error: '#VALUE!' }, `${name} with one argument`);
  }
  for (const name of ['HARMEAN', 'AVEDEV', 'DEVSQ', 'STDEVA', 'STDEVPA', 'VARA', 'VARPA']) {
    const spec = FUNCTIONS.get(name);
    assert.ok(spec, `${name} is registered`);
    assert.deepEqual(spec.call([], context), { error: '#VALUE!' }, `${name} without arguments`);
  }
});

// ---------------------------------------------------------------------------
// Case of letters

function withLocale(locale: FormatLocale | undefined, cells: Cells = {}): CalcEngine {
  return engineOf(cells, { ...NO_LIMIT, locale: () => locale });
}

test('UPPER, LOWER and PROPER follow the Turkish letters in a Turkish locale', () => {
  const turkish = withLocale(LOCALE_TR);
  assert.equal(calc('UPPER("istanbul ılık")', turkish), 'İSTANBUL ILIK');
  assert.equal(calc('LOWER("İSTANBUL ILIK")', turkish), 'istanbul ılık');
  assert.equal(calc('PROPER("istanbul ığdır IŞIK İZMİR")', turkish), 'İstanbul Iğdır Işık İzmir');
  assert.equal(calc('UPPER("iıİI")', turkish), 'İIİI');
  assert.equal(calc('LOWER("iıİI")', turkish), 'iıiı');
  // The other letters go the usual way.
  assert.equal(calc('UPPER("çğöşü abc")', turkish), 'ÇĞÖŞÜ ABC');
  assert.equal(calc('LOWER("ÇĞÖŞÜ ABC")', turkish), 'çğöşü abc');
  // A capital I with a combining dot after it is one dotted capital.
  assert.equal(calc('LOWER("I\u0307Z")', turkish), 'iz');
  assert.equal(calc('UPPER("")', turkish), '');
  assert.equal(calc('LOWER(123)', turkish), '123');
  assert.equal(calc('UPPER(TRUE)', turkish), 'TRUE');
  assert.deepEqual(calc('UPPER(#N/A)', turkish), { error: '#N/A' });
  assert.deepEqual(calc('PROPER(1/0)', turkish), { error: '#DIV/0!' });
});

test('without a Turkish locale the case of letters is the language-neutral one', () => {
  for (const engine of [blank, withLocale(undefined), withLocale(LOCALE_EN)]) {
    assert.equal(calc('UPPER("istanbul ılık")', engine), 'ISTANBUL ILIK');
    assert.equal(calc('LOWER("ISTANBUL")', engine), 'istanbul');
    assert.equal(calc('PROPER("istanbul ığdır")', engine), 'Istanbul Iğdır');
    assert.equal(calc('UPPER("çğöşü")', engine), 'ÇĞÖŞÜ');
  }
});

test('the locale is asked when a formula is calculated: the next calculation follows a change', () => {
  let locale: FormatLocale = LOCALE_EN;
  const workbook = book({ A1: 'izmir', B1: '=UPPER(A1)', B2: '=LOWER("IŞIK")', B3: '=PROPER(A1)' });
  const engine = new CalcEngine(workbook, { ...NO_LIMIT, locale: () => locale });
  const value = (address: string) => {
    const at = parseCellRef(address) as { row: number; col: number };
    return getCell(workbook.sheets[0], at.row, at.col)!.value;
  };
  engine.recalculate(undefined, { overwrite: true });
  assert.deepEqual([value('B1'), value('B2'), value('B3')], ['IZMIR', 'işik', 'Izmir']);
  locale = LOCALE_TR;
  engine.recalculate(undefined, { overwrite: true });
  assert.deepEqual([value('B1'), value('B2'), value('B3')], ['İZMİR', 'ışık', 'İzmir']);
});

test('functions that compare text do not depend on the locale', () => {
  const cells: Cells = { A1: 'IŞIK', A2: 'ışık', A3: 'İzmir', A4: 'izmir', B1: 1, B2: 2, B3: 3, B4: 4 };
  const formulas = ['MATCH("ışık",A1:A4,0)', 'COUNTIF(A1:A4,"izmir")', 'VLOOKUP("İZMİR",A1:B4,2,FALSE)', 'EXACT("i","İ")', 'SEARCH("I","ai")', 'XLOOKUP("izmir",A1:A4,B1:B4,0)'];
  const english = withLocale(LOCALE_EN, cells);
  const turkish = withLocale(LOCALE_TR, cells);
  for (const formula of formulas) assert.deepEqual(calc(formula, turkish), calc(formula, english), formula);
});

// ---------------------------------------------------------------------------
// Turkish names

/** The pairs of the table in formula-locale.ts, read from what the translation does with each name. */
function storedName(local: string): string {
  const stored = formulaFromLocal(`=${local}()`, LOCALE_TR);
  return stored.slice(1, stored.length - 2);
}

test('every function typed by its Turkish name is one the engine has', () => {
  const names: [string, string][] = [
    ['ÜS', 'EXP'], ['İŞARET', 'SIGN'], ['ÇARPINIM', 'FACT'], ['ÇİFTFAKTÖR', 'FACTDOUBLE'], ['KOMBİNASYON', 'COMBIN'], ['PERMÜTASYON', 'PERMUT'],
    ['OBEB', 'GCD'], ['OKEK', 'LCM'], ['KYUVARLA', 'MROUND'], ['BÖLÜM', 'QUOTIENT'], ['ÇİFT', 'EVEN'], ['TEK', 'ODD'], ['DERECE', 'DEGREES'],
    ['RADYAN', 'RADIANS'], ['TOPKARE', 'SUMSQ'], ['KAREKÖKPİ', 'SQRTPI'], ['ROMEN', 'ROMAN'], ['TAVANAYUVARLA.MATEMATİK', 'CEILING.MATH'],
    ['TABANAYUVARLA.MATEMATİK', 'FLOOR.MATH'], ['TAVANAYUVARLA.DUYARLI', 'CEILING.PRECISE'], ['TABANAYUVARLA.DUYARLI', 'FLOOR.PRECISE'],
    ['ISO.TAVAN', 'ISO.CEILING'], ['TOPLAMA', 'AGGREGATE'], ['DÇARP', 'MMULT'], ['TOPX2EY2', 'SUMX2MY2'], ['TOPX2AY2', 'SUMX2PY2'], ['TOPXEY2', 'SUMXMY2'],
    ['ORTALAMAA', 'AVERAGEA'], ['MAKA', 'MAXA'], ['MİNA', 'MINA'], ['ÇOKEĞERMAK', 'MAXIFS'], ['ÇOKEĞERMİN', 'MINIFS'], ['ENÇOK_OLAN', 'MODE'],
    ['ENÇOK_OLAN.TEK', 'MODE.SNGL'], ['STDSAPMA', 'STDEV'], ['STDSAPMA.S', 'STDEV.S'], ['STDSAPMA.P', 'STDEV.P'], ['STDSAPMAS', 'STDEVP'],
    ['STDSAPMAA', 'STDEVA'], ['STDSAPMASA', 'STDEVPA'], ['VARS', 'VARP'], ['VARSA', 'VARPA'], ['RANK.EŞİT', 'RANK.EQ'], ['RANK.ORT', 'RANK.AVG'],
    ['YÜZDEBİRLİK', 'PERCENTILE'], ['YÜZDEBİRLİK.DHL', 'PERCENTILE.INC'], ['YÜZDEBİRLİK.HRC', 'PERCENTILE.EXC'], ['DÖRTTEBİRLİK', 'QUARTILE'],
    ['DÖRTTEBİRLİK.DHL', 'QUARTILE.INC'], ['DÖRTTEBİRLİK.HRC', 'QUARTILE.EXC'], ['KORELASYON', 'CORREL'], ['EĞİM', 'SLOPE'], ['KESMENOKTASI', 'INTERCEPT'],
    ['TAHMİN', 'FORECAST'], ['TAHMİN.DOĞRUSAL', 'FORECAST.LINEAR'], ['GEOORT', 'GEOMEAN'], ['HARORT', 'HARMEAN'], ['KIRPORTALAMA', 'TRIMMEAN'],
    ['SIKLIK', 'FREQUENCY'], ['ORTSAP', 'AVEDEV'], ['SAPKARE', 'DEVSQ'], ['RKARE', 'RSQ'], ['KOVARYANS', 'COVAR'], ['KOVARYANS.P', 'COVARIANCE.P'],
    ['KOVARYANS.S', 'COVARIANCE.S'], ['NORM.DAĞ', 'NORM.DIST'], ['NORMDAĞ', 'NORMDIST'], ['NORM.TERS', 'NORM.INV'], ['NORMTERS', 'NORMINV'],
    ['NORM.S.DAĞ', 'NORM.S.DIST'], ['NORMSDAĞ', 'NORMSDIST'], ['NORM.S.TERS', 'NORM.S.INV'], ['NORMSTERS', 'NORMSINV'], ['YAZIM.DÜZENİ', 'PROPER'],
    ['TEMİZ', 'CLEAN'], ['KOD', 'CODE'], ['SAYIDÜZENLE', 'FIXED'], ['LİRA', 'DOLLAR'], ['SAYIDEĞERİ', 'NUMBERVALUE'], ['ARALIKBİRLEŞTİR', 'CONCAT'],
    ['METİNBİRLEŞTİR', 'TEXTJOIN'], ['M', 'T'], ['S', 'N'], ['TARİHSAYISI', 'DATEVALUE'], ['ZAMANSAYISI', 'TIMEVALUE'], ['GÜNSAY', 'DAYS'],
    ['GÜN360', 'DAYS360'], ['HAFTASAY', 'WEEKNUM'], ['ISOHAFTASAY', 'ISOWEEKNUM'], ['YILORAN', 'YEARFRAC'], ['ETARİHLİ', 'DATEDIF'],
    ['TAMİŞGÜNÜ.ULUSL', 'NETWORKDAYS.INTL'], ['İŞGÜNÜ.ULUSL', 'WORKDAY.INTL'], ['ELEMAN', 'CHOOSE'], ['ADRES', 'ADDRESS'], ['ALANSAY', 'AREAS'],
    ['DEVRİK_DÖNÜŞÜM', 'TRANSPOSE'], ['ÇAPRAZARA', 'XLOOKUP'], ['ÇAPRAZEŞLEŞTİR', 'XMATCH'], ['KÖPRÜ', 'HYPERLINK'], ['FORMÜLMETNİ', 'FORMULATEXT'],
    ['ÇOKEĞER', 'IFS'], ['İLKEŞLEŞEN', 'SWITCH'], ['ÖZELVEYA', 'XOR'], ['YOKSAY', 'NA'], ['EHATA', 'ISERR'], ['EYOKSA', 'ISNA'],
    ['EMETİNDEĞİLSE', 'ISNONTEXT'], ['EMANTIKSALSA', 'ISLOGICAL'], ['ÇİFTMİ', 'ISEVEN'], ['TEKMİ', 'ISODD'], ['EREFSE', 'ISREF'], ['EFORMÜLSE', 'ISFORMULA'],
    ['HATA.TİPİ', 'ERROR.TYPE'], ['TİP', 'TYPE'], ['SAYFA', 'SHEET'], ['SAYFALAR', 'SHEETS'], ['HÜCRE', 'CELL'], ['DEVRESEL_ÖDEME', 'PMT'], ['BD', 'PV'],
    ['GD', 'FV'], ['TAKSİT_SAYISI', 'NPER'], ['FAİZ_ORANI', 'RATE'], ['NBD', 'NPV'], ['İÇ_VERİM_ORANI', 'IRR'], ['ANBD', 'XNPV'], ['AİÇVERİMORANI', 'XIRR'],
    ['FAİZTUTARI', 'IPMT'], ['ANA_PARA_ÖDEMESİ', 'PPMT'], ['DA', 'SLN'], ['AZALANBAKİYE', 'DB'], ['ÇİFTAZALANBAKİYE', 'DDB'], ['YAT', 'SYD'], ['ETKİN', 'EFFECT'],
    ['SİN', 'SIN'], ['ASİN', 'ASIN'], ['SİNH', 'SINH'], ['ASİNH', 'ASINH'], ['Pİ', 'PI'], ['NOMİNAL', 'NOMINAL'],
    ['VSEÇTOPLA', 'DSUM'], ['VSEÇSAY', 'DCOUNT'], ['VSEÇSAYDOLU', 'DCOUNTA'], ['VSEÇORT', 'DAVERAGE'], ['VSEÇMAK', 'DMAX'], ['VSEÇMİN', 'DMIN'], ['VAL', 'DGET'],
    // Names that were there before.
    ['TOPLA', 'SUM'], ['DÜŞEYARA', 'VLOOKUP'], ['ÇOKETOPLA', 'SUMIFS'], ['SERİTARİH', 'EDATE'], ['BÜYÜKHARF', 'UPPER'],
  ];
  for (const [local, stored] of names) {
    assert.equal(storedName(local), stored, local);
    assert.ok(FUNCTIONS.has(stored) || isBuiltinForm(stored), `${stored} (${local}) is a function of the engine`);
    // Small letters, as a Turkish keyboard writes them, are the same name.
    const small = local.replace(/I/g, 'ı').replace(/İ/g, 'i').toLowerCase();
    assert.equal(storedName(small), stored, `${local} in small letters (${small})`);
    // A Turkish name is never the English name of another function.
    const key = local.replace(/[iıİ]/g, 'I').toUpperCase();
    assert.ok(key === stored || !(FUNCTIONS.has(key) || isBuiltinForm(key)), `${local} is not the English name of another function`);
  }
});

test('an English name stays what it is in a Turkish formula, and a name is only a function where it is called', () => {
  assert.equal(formulaFromLocal('=HARMEAN(1;2)', LOCALE_TR), '=HARMEAN(1,2)');
  assert.equal(formulaFromLocal('=LN(1,5)+cos(0)', LOCALE_TR), '=LN(1.5)+cos(0)');
  // "S" and "M" are columns, and "DA" may be a defined name, wherever no bracket follows.
  assert.equal(formulaFromLocal('=S1+M2+DA', LOCALE_TR), '=S1+M2+DA');
  assert.equal(formulaFromLocal('=S(S1)&M(M2)', LOCALE_TR), '=N(S1)&T(M2)');
  assert.equal(formulaFromLocal('=TOPLA(S:S)', LOCALE_TR), '=SUM(S:S)');
});

test('SİN and the other names that only differ by a dotted İ are stored by their English names', () => {
  assert.equal(formulaFromLocal('=SİN(Pİ()/2)+sin(0)', LOCALE_TR), '=SIN(PI()/2)+SIN(0)');
  near(calc(formulaFromLocal('=ASİNH(1)', LOCALE_TR).slice(1)), 0.881373587019543, 'ASİNH(1)');
  near(calc(formulaFromLocal('=SİN(0,5)', LOCALE_TR).slice(1)), Math.sin(0.5), 'SİN(0,5)');
});

// ---------------------------------------------------------------------------
// Every family, typed the Turkish way

const NUMBERS = [12, 7, 3, 25, 7, 18, 30, 4, 9, 15];
const FRUIT = ['elma', 'armut', 'elma', 'kiraz', 'armut', 'elma', 'kiraz', 'elma', 'armut', 'kiraz'];
/** 15 January 2024 and every twentieth day after it. */
const DATES = [45306, 45326, 45346, 45366, 45386, 45406, 45426, 45446, 45466, 45486];

function table(): Cells {
  const cells: Cells = {};
  for (let i = 0; i < 10; i++) {
    cells[`A${i + 1}`] = NUMBERS[i];
    cells[`B${i + 1}`] = FRUIT[i];
    cells[`C${i + 1}`] = DATES[i];
  }
  return cells;
}

type Expected = number | string | boolean | { error: string } | { near: number };

function typed(cases: [string, Expected][], cells: Cells = table()): void {
  const engine = withLocale(LOCALE_TR, cells);
  for (const [local, expected] of cases) {
    const stored = formulaFromLocal(`=${local}`, LOCALE_TR).slice(1);
    const actual = calc(stored, engine);
    if (expected !== null && typeof expected === 'object' && 'near' in expected) near(actual, expected.near, local);
    else assert.deepEqual(actual, expected, `${local} (stored as ${stored})`);
  }
}

test('typed in Turkish: conditional totals', () => {
  typed([
    ['ÇOKETOPLA(A1:A10;B1:B10;"elma";A1:A10;">=5")', 30],
    ['ÇOKEĞERSAY(B1:B10;"k*";A1:A10;"<>25")', 2],
    ['ÇOKEĞERORTALAMA(A1:A10;B1:B10;"armut")', { near: 7.666666666666667 }],
    ['ÇOKEĞERMAK(A1:A10;B1:B10;"elma")', 18],
    ['ÇOKEĞERMİN(A1:A10;B1:B10;"elma")', 3],
    ['ÇOKEĞERMİN(A1:A10;B1:B10;"muz")', 0],
    ['ETOPLA(B1:B10;"?iraz";A1:A10)', 70],
    ['EĞERSAY(A1:A10;">7")', 6],
    ['EĞERSAY(A1:A10;">"&7,5)', 6],
    ['EĞERSAY(B1:B12;"")', 2],
    ['EĞERSAY(B1:B12;"<>")', 10],
    ['EĞERSAY(B1:B12;"=")', 2],
    ['TOPLA.ÇARPIM((B1:B10="elma")*A1:A10)', 37],
    ['ALTTOPLAM(109;A1:A10)', 130],
    ['ALTTOPLAM(1;A1:A10)', 13],
    ['TOPLAMA(4;6;A1:A10)', 30],
    ['TOPLAMA(14;6;A1:A10;2)', 25],
    ['ÇOKEĞERORTALAMA(A1:A10;B1:B10;"muz")', { error: '#DIV/0!' }],
    ['ÇOKETOPLA(A1:A10;B1:B9;"elma")', { error: '#VALUE!' }],
  ]);
});

test('typed in Turkish: lookup and reference', () => {
  typed([
    ['DÜŞEYARA("kiraz";B1:C10;2;YANLIŞ)', 45366],
    ['DÜŞEYARA("KİRAZ";B1:C10;2;YANLIŞ)', { error: '#N/A' }],
    ['YATAYARA(12;A1:C2;2;YANLIŞ)', 7],
    ['ARA(16;{3;7;12;18;25};{"a";"b";"c";"d";"e"})', 'c'],
    ['ÇAPRAZARA(18;A1:A10;B1:B10)', 'elma'],
    ['ÇAPRAZARA(8;A1:A10;B1:B10;"yok")', 'yok'],
    ['ÇAPRAZARA(8;A1:A10;B1:B10;"yok";-1)', 'armut'],
    ['ÇAPRAZARA(8;A1:A10;B1:B10;"yok";1)', 'armut'],
    ['ÇAPRAZARA("a*";B1:B10;A1:A10;;2;-1)', 9],
    ['ÇAPRAZEŞLEŞTİR(30;A1:A10)', 7],
    ['ÇAPRAZEŞLEŞTİR(7;A1:A10;0;-1)', 5],
    ['KAÇINCI(9;A1:A10;0)', 9],
    ['İNDİS(A1:B10;4;2)', 'kiraz'],
    ['İNDİS((A1:A10;B1:B10);2;1;2)', 'armut'],
    ['ELEMAN(2;"a";"b";"c")', 'b'],
    ['ELEMAN(4;"a";"b";"c")', { error: '#VALUE!' }],
    ['ADRES(2;3)', '$C$2'],
    ['ADRES(2;3;4)', 'C2'],
    ['ALANSAY((A1:A2;B1:B2))', 2],
    ['SATIR(B7)', 7],
    ['SÜTUN(C1)', 3],
    ['SATIRSAY(A1:B10)', 10],
    ['SÜTUNSAY(A1:B10)', 2],
    ['TOPLA(KAYDIR(A1;1;0;3;1))', 35],
    ['DOLAYLI("A"&4)', 25],
    ['DOLAYLI("R4C1";YANLIŞ)', { error: '#REF!' }],
    ['TOPLA(DEVRİK_DÖNÜŞÜM(A1:A3))', 22],
  ]);
});

test('typed in Turkish: text', () => {
  typed([
    ['METİNBİRLEŞTİR("-";DOĞRU;B1:B3)', 'elma-armut-elma'],
    ['ARALIKBİRLEŞTİR("a";1,5;B4)', 'a1.5kiraz'],
    ['YERİNEKOY("a-b-c";"-";"+";2)', 'a-b+c'],
    ['DEĞİŞTİR("abcdef";2;3;"X")', 'aXef'],
    ['YİNELE("ab";3)', 'ababab'],
    ['YAZIM.DÜZENİ("ışık izmir")', 'Işık İzmir'],
    ['BÜYÜKHARF(B1)', 'ELMA'],
    ['BÜYÜKHARF("kiraz")', 'KİRAZ'],
    ['KÜÇÜKHARF("KIRIK")', 'kırık'],
    ['TEMİZ("a"&DAMGA(7)&"b")', 'ab'],
    ['KIRP("  a   b  ")', 'a b'],
    ['KOD("A")', 65],
    ['ÖZDEŞ("a";"A")', false],
    ['BUL("r";"armut")', 2],
    ['MBUL("M?T";"armut")', 3],
    ['SAYIYAÇEVİR("12")+1', 13],
    ['SAYIDEĞERİ("1.234,5";",";".")', 1234.5],
    ['SAYIDÜZENLE(1234,567;1;DOĞRU)', '1234.6'],
    ['METNEÇEVİR(1234,5;"0.00")', '1234.50'],
    ['M(5)', ''],
    ['M("a")', 'a'],
    ['S(DOĞRU)', 1],
    ['S("a")', 0],
    ['UZUNLUK(YİNELE("x";40000))', { error: '#VALUE!' }],
    ['BUL("z";"armut")', { error: '#VALUE!' }],
  ]);
});

test('typed in Turkish: dates', () => {
  typed([
    ['SERİTARİH(TARİH(2024;1;31);1)', 45351],
    ['SERİAY(TARİH(2024;1;15);0)', 45322],
    ['ETARİHLİ(TARİH(2020;2;29);TARİH(2024;2;28);"y")', 3],
    ['ETARİHLİ(C1;C10;"d")', 180],
    ['TAMİŞGÜNÜ(TARİH(2024;1;1);TARİH(2024;1;31))', 23],
    ['TAMİŞGÜNÜ.ULUSL(TARİH(2024;1;1);TARİH(2024;1;31);"0000011")', 23],
    ['TAMİŞGÜNÜ.ULUSL(TARİH(2024;1;1);TARİH(2024;1;31);11)', 27],
    ['İŞGÜNÜ(TARİH(2024;1;1);10)', 45306],
    ['İŞGÜNÜ.ULUSL(TARİH(2024;1;1);10;1)', 45306],
    ['HAFTASAY(C4)', 11],
    ['ISOHAFTASAY(C4)', 11],
    ['HAFTANINGÜNÜ(C4)', 6],
    ['HAFTANINGÜNÜ(C4;2)', 5],
    ['HAFTANINGÜNÜ(C4;3)', 4],
    ['HAFTANINGÜNÜ(C4;11)', 5],
    ['HAFTANINGÜNÜ(C4;16)', 7],
    ['HAFTANINGÜNÜ(C4;17)', 6],
    ['HAFTANINGÜNÜ(C4;4)', { error: '#NUM!' }],
    ['YILORAN(TARİH(2024;1;1);TARİH(2024;7;1);1)', { near: 0.4972677595628415 }],
    ['YILORAN(TARİH(2024;1;1);TARİH(2024;7;1);0)', 0.5],
    ['YILORAN(TARİH(2024;1;1);TARİH(2024;7;1);2)', { near: 182 / 360 }],
    ['YILORAN(TARİH(2024;1;1);TARİH(2024;7;1);3)', { near: 182 / 365 }],
    ['YILORAN(TARİH(2024;1;1);TARİH(2024;7;1);4)', 0.5],
    ['YILORAN(TARİH(2024;1;1);TARİH(2024;7;1);5)', { error: '#NUM!' }],
    ['GÜN360(TARİH(2024;1;31);TARİH(2024;3;31))', 60],
    ['GÜNSAY(TARİH(2024;3;1);TARİH(2024;2;1))', 29],
    ['TARİHSAYISI("2024-03-15")', 45366],
    ['ZAMANSAYISI("6:00")', 0.25],
    ['ZAMAN(6;0;0)', 0.25],
    // The leap day 1900 never was, and has a serial number all the same.
    ['TARİH(1900;2;29)', 60],
    ['GÜN(60)', 29],
    ['AY(61)', 3],
    ['SERİTARİH("abc";1)', { error: '#VALUE!' }],
  ]);
});

test('dates in the 1904 system are 1,462 days earlier', () => {
  const engine = new CalcEngine(book({}, true), NO_LIMIT);
  assert.equal(calc('DATE(2024,1,31)', engine), 45322 - 1462);
  assert.equal(calc('EDATE(DATE(2024,1,31),1)', engine), 45351 - 1462);
  assert.equal(calc('EOMONTH(DATE(2024,1,15),0)', engine), 45322 - 1462);
  assert.equal(calc('WEEKDAY(DATE(2024,3,15))', engine), 6);
  assert.equal(calc('NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31))', engine), 23);
  assert.equal(calc('ISOWEEKNUM(DATE(2024,3,15))', engine), 11);
  assert.equal(calc('YEAR(0)', engine), 1904);
});

test('typed in Turkish: mathematics and statistics', () => {
  typed([
    ['KYUVARLA(10;3)', 9],
    ['YUKARIYUVARLA(3,14159;2)', 3.15],
    ['AŞAĞIYUVARLA(-3,14159;2)', -3.14],
    ['TAVANAYUVARLA.MATEMATİK(-5,5;2)', -4],
    ['TABANAYUVARLA.DUYARLI(-5,5;2)', -6],
    ['NSAT(-1,5)', -1],
    ['TAMSAYI(-1,5)', -2],
    ['BÖLÜM(7;2)', 3],
    ['OBEB(12;18)', 6],
    ['OKEK(4;6)', 12],
    ['ÇARPINIM(5)', 120],
    ['KOMBİNASYON(5;2)', 10],
    ['PERMÜTASYON(5;2)', 20],
    ['TOPKARE(3;4)', 25],
    ['İŞARET(-3)', -1],
    ['ÜS(0)', 1],
    ['ORTANCA(A1:A10)', 10.5],
    ['ENÇOK_OLAN(A1:A10)', 7],
    ['BÜYÜK(A1:A10;2)', 25],
    ['KÜÇÜK(A1:A10;2)', 4],
    ['RANK.ORT(7;A1:A10)', 7.5],
    ['RANK.EŞİT(7;A1:A10)', 7],
    ['YÜZDEBİRLİK.DHL(A1:A10;0,9)', { near: 25.5 }],
    ['DÖRTTEBİRLİK.HRC(A1:A10;1)', { near: 6.25 }],
    ['STDSAPMA.S(A1:A10)', { near: 9.018499505645789 }],
    ['VARS(A1:A10)', { near: 73.2 }],
    ['ORTALAMAA(A1:A2;B1)', { near: 19 / 3 }],
    ['BAĞ_DEĞ_DOLU_SAY(A1:B11)', 20],
    ['BOŞLUKSAY(A1:B11)', 2],
    ['MAKA(B1:B3)', 0],
    ['KORELASYON(A1:A10;C1:C10)', { near: 0.14649405007169522 }],
    ['EĞİM(A1:A10;C1:C10)', { near: 0.02181818181818182 }],
    ['KESMENOKTASI(A1:A10;C1:C10)', { near: -977.4581818181819 }],
    ['TAHMİN.DOĞRUSAL(45500;A1:A10;C1:C10)', { near: 15.269090909091005 }],
    ['GEOORT(A1:A10)', { near: 10.254691765285898 }],
    ['HARORT(A1:A10)', { near: 7.942511346444781 }],
    ['ORTSAP(1;2;3)', { near: 2 / 3 }],
    ['SAPKARE(1;2;3)', 2],
    ['KAREKÖK(-1)', { error: '#NUM!' }],
    ['MOD(5;0)', { error: '#DIV/0!' }],
    ['ORTANCA(B1:B10)', { error: '#NUM!' }],
  ]);
});

test('typed in Turkish: logic and information', () => {
  typed([
    ['ÇOKEĞER(A1>20;"b";A1>10;"o";DOĞRU;"k")', 'o'],
    ['ÇOKEĞER(A1>20;"b")', { error: '#N/A' }],
    ['İLKEŞLEŞEN(B1;"armut";1;"elma";2;0)', 2],
    ['İLKEŞLEŞEN(B4;"armut";1;"elma";2)', { error: '#N/A' }],
    ['ÖZELVEYA(DOĞRU;DOĞRU;DOĞRU)', true],
    ['ÖZELVEYA(DOĞRU;DOĞRU)', false],
    ['EĞERYOKSA(KAÇINCI(99;A1:A10;0);"yok")', 'yok'],
    ['EĞERHATA(1/0;"hata")', 'hata'],
    ['EĞERYOKSA(1/0;"yok")', { error: '#DIV/0!' }],
    ['EBOŞSA(A11)', true],
    ['ESAYIYSA(A1)', true],
    ['EMETİNSE(A1)', false],
    ['EMETİNDEĞİLSE(A11)', true],
    ['EMANTIKSALSA(YANLIŞ)', true],
    ['EHATALIYSA(#YOK)', true],
    ['EHATA(#YOK)', false],
    ['EYOKSA(YOKSAY())', true],
    ['EFORMÜLSE(A1)', false],
    ['EREFSE(A1)', true],
    ['ÇİFTMİ(4)', true],
    ['TEKMİ(4)', false],
    ['HATA.TİPİ(#SAYI/0!)', 2],
    ['HATA.TİPİ(1)', { error: '#N/A' }],
    ['TİP("a")', 2],
    ['TİP(A1)', 1],
    ['SAYFA()', 1],
    ['SAYFALAR()', 1],
  ]);
});

test('typed in Turkish: finance', () => {
  typed([
    // Excel's documentation gives these, rounded: -1,037.03, 2,581.40, -59,777.15, 59.6738657, 1,188.44, 8.7 %, 0.77 %.
    ['DEVRESEL_ÖDEME(0,08/12;10;10000)', { near: -1037.0320893591636 }],
    ['GD(0,06/12;10;-200;-500;1)', { near: 2581.4033740601362 }],
    ['BD(0,08/12;12*20;500;;0)', { near: -59777.14585118777 }],
    ['TAKSİT_SAYISI(0,12/12;-100;-1000;10000;1)', { near: 59.67386567429457 }],
    ['NBD(0,1;-10000;3000;4200;6800)', { near: 1188.4434123352212 }],
    ['DA(30000;7500;10)', 2250],
    ['ÇİFTAZALANBAKİYE(2400;300;10;1)', 480],
    ['AZALANBAKİYE(1000000;100000;6;1;7)', { near: 186083.33333333334 }],
    ['FAİZTUTARI(0,1;1;3;8000)+ANA_PARA_ÖDEMESİ(0,1;1;3;8000)-DEVRESEL_ÖDEME(0,1;3;8000)', { near: 0 }],
    ['FAİZ_ORANI(10;100;100)', { error: '#NUM!' }],
    ['DEVRESEL_ÖDEME("a";10;1)', { error: '#VALUE!' }],
  ]);
  const engine = withLocale(LOCALE_TR);
  const rate = calc(formulaFromLocal('=FAİZ_ORANI(4*12;-200;8000)', LOCALE_TR).slice(1), engine) as number;
  assert.ok(Math.abs(rate - 0.007701472488202) < 1e-9, `RATE gives ${rate}`);
  const irr = calc(formulaFromLocal('=İÇ_VERİM_ORANI({-70000;12000;15000;18000;21000;26000})', LOCALE_TR).slice(1), engine) as number;
  assert.ok(Math.abs(irr - 0.0866309480365316) < 1e-9, `IRR gives ${irr}`);
  // XIRR and XNPV agree with each other: the rate XIRR finds makes XNPV nothing.
  const flows = '{-10000,2750,4250,3250,2750},{39448,39508,39751,39859,39904}';
  const xirr = calc(`XIRR(${flows})`, engine) as number;
  // Excel's documentation: 0.373362535.
  assert.ok(Math.abs(xirr - 0.373362535) < 1e-8, `XIRR gives ${xirr}`);
  const left = calc(`XNPV(${xirr},${flows})`, engine) as number;
  assert.ok(Math.abs(left) < 1e-4, `XNPV at the rate of XIRR gives ${left}`);
});

test('a Turkish formula with a comma that is no decimal is refused', () => {
  assert.throws(() => formulaFromLocal('=HARORT(4,5,8)', LOCALE_TR), RangeError);
  assert.equal(formulaFromLocal('=HARORT(4,5;8)', LOCALE_TR), '=HARMEAN(4.5,8)');
});

// ---------------------------------------------------------------------------
// Volatile functions

test('RAND and RANDBETWEEN are calculated again by every recalculation', () => {
  let next = 0;
  const workbook = book({ A1: '=RAND()', A2: '=RANDBETWEEN(1,100)', A3: '=A1*0+5', B1: 1 });
  const engine = new CalcEngine(workbook, { ...NO_LIMIT, random: () => (next++ % 10) / 10 });
  const value = (row: number) => getCell(workbook.sheets[0], row, 0)!.value;
  engine.recalculate();
  const before = [value(0), value(1)];
  assert.equal(value(2), 5);
  // A cell nothing depends on changes: the volatile formulas are calculated all the same.
  engine.recalculate([{ sheet: 0, row: 0, col: 1 }]);
  const after = [value(0), value(1)];
  assert.notDeepEqual(after, before, 'new random numbers');
  for (const v of after) assert.equal(typeof v, 'number', 'a random number is a number');
  assert.ok((after[1] as number) >= 1 && (after[1] as number) <= 100, `RANDBETWEEN(1,100) gives ${after[1]}`);
  assert.deepEqual(calc('RANDBETWEEN(5,1)', engine), { error: '#NUM!' });
});

// ---------------------------------------------------------------------------
// INDEX, reference form

test('INDEX picks one of several areas', () => {
  const engine = engineOf({ A1: 1, A2: 2, A3: 3, B1: 'a', B2: 'b', B3: 'c', D1: 10, D2: 20, E1: 30, E2: 40 });
  assert.equal(calc('INDEX((A1:B3,D1:E2),2,2)', engine), 'b', 'the first area when none is named');
  assert.equal(calc('INDEX((A1:B3,D1:E2),2,2,1)', engine), 'b');
  assert.equal(calc('INDEX((A1:B3,D1:E2),2,2,2)', engine), 40);
  assert.equal(calc('INDEX((A1:A3,B1:B3,D1:E2),3,1,2)', engine), 'c');
  assert.equal(calc('SUM(INDEX((A1:B3,D1:E2),0,0,2))', engine), 100, 'row and column 0 give the whole area');
  assert.equal(calc('SUM(INDEX((A1:B3,D1:E2),0,1,2))', engine), 30, 'a whole column of the area');
  assert.equal(calc('ROWS(INDEX((A1:B3,D1:E2),0,1,2))', engine), 2, 'what comes back is a reference');
  isErrorOf(calc('INDEX((A1:B3,D1:E2),1,1,3)', engine), '#REF!', 'more than there are areas');
  isErrorOf(calc('INDEX((A1:B3,D1:E2),1,1,0)', engine), '#VALUE!', 'area 0');
  isErrorOf(calc('INDEX((A1:B3,D1:E2),3,1,2)', engine), '#REF!', 'a row the area does not have');
  isErrorOf(calc('INDEX((A1:B3,D1:E2),1,1,"x")', engine), '#VALUE!', 'an area that is no number');
  isErrorOf(calc('INDEX((A1:B3,D1:E2),1,1,#N/A)', engine), '#N/A', 'an error as the area');
  // One area: the array form and the reference form are the same.
  assert.equal(calc('INDEX(A1:B3,3,2,1)', engine), 'c');
  isErrorOf(calc('INDEX(A1:B3,1,1,2)', engine), '#REF!', 'a second area of one');
});

test('a formula with INDEX over several areas follows the cells of each area', () => {
  const workbook = book({ A1: 1, A2: 2, D1: 10, D2: 20, G1: '=INDEX((A1:A2,D1:D2),2,1,2)', G2: '=INDEX((A1:A2,D1:D2),1,1,H1)', H1: 1 });
  const engine = new CalcEngine(workbook, NO_LIMIT);
  const value = (address: string) => {
    const at = parseCellRef(address) as { row: number; col: number };
    return getCell(workbook.sheets[0], at.row, at.col)!.value;
  };
  const set = (address: string, to: number) => {
    const at = parseCellRef(address) as { row: number; col: number };
    const row = workbook.sheets[0].rows.get(at.row)!;
    const cells = row.cells.slice();
    cells[at.col] = makeCell(to);
    workbook.sheets[0].rows.set(at.row, { ...row, cells });
    return { sheet: 0, row: at.row, col: at.col };
  };
  engine.recalculate();
  assert.deepEqual([value('G1'), value('G2')], [20, 1]);
  assert.deepEqual(engine.unsupported, [], 'nothing is left to the value of the file');
  engine.recalculate([set('D2', 21)]);
  assert.equal(value('G1'), 21, 'a cell of the second area changed');
  engine.recalculate([set('H1', 2)]);
  assert.equal(value('G2'), 10, 'the number of the area changed');
});

// ---------------------------------------------------------------------------
// After the review

test('INDEX takes areas however the union is bracketed', () => {
  const engine = engineOf({ A1: 1, A2: 2, A3: 3, B1: 10, B2: 20, B3: 30, C1: 100, C2: 200, C3: 300 });
  assert.equal(calc('INDEX((A1:A3,(B1:B3,C1:C3)),2,1,3)', engine), 200);
  assert.equal(calc('INDEX((A1:A3,(B1:B3,C1:C3)),2,1,2)', engine), 20);
  assert.equal(calc('INDEX(((A1:A3,B1:B3),C1:C3),2,1,1)', engine), 2);
  assert.equal(calc('INDEX(((A1:A3,B1:B3),(C1:C3,A1:A3)),3,1,4)', engine), 3);
  isErrorOf(calc('INDEX((A1:A3,(B1:B3,C1:C3)),2,1,4)', engine), '#REF!', 'a fourth area of three');
  assert.equal(calc('AREAS((A1:A3,(B1:B3,C1:C3)))', engine), 3);
  assert.equal(calc('SUM((A1:A3,(B1:B3,C1:C3)))', engine), 666);
});

test('a formula with a union inside a union is calculated and follows every area', () => {
  const workbook = book({ A1: 1, B1: 10, C1: 100, C2: 200, G1: '=INDEX((A1:A2,(B1:B2,C1:C2)),2,1,3)', G2: '=AREAS((A1:A2,(B1:B2,C1:C2)))', G3: '=SUM((A1:A2,(B1:B2,C1:C2)))' });
  const engine = new CalcEngine(workbook, NO_LIMIT);
  const value = (row: number) => getCell(workbook.sheets[0], row, 6)!.value;
  engine.recalculate();
  assert.deepEqual(engine.unsupported, [], 'nothing is left to the value of the file');
  assert.deepEqual([value(0), value(1), value(2)], [200, 3, 311]);
  const row = workbook.sheets[0].rows.get(1)!;
  const cells = row.cells.slice();
  cells[2] = makeCell(201);
  workbook.sheets[0].rows.set(1, { ...row, cells });
  engine.recalculate([{ sheet: 0, row: 1, col: 2 }]);
  assert.deepEqual([value(0), value(2)], [201, 312], 'a cell of the innermost area changed');
});

test('a combining dot belongs to its letter: UPPER, LOWER and PROPER in Turkish', () => {
  const turkish = withLocale(LOCALE_TR);
  assert.equal(calc('UPPER("i\u0307stanbul")', turkish), 'İSTANBUL', 'no second dot on the dotted capital');
  assert.equal(calc('UPPER("ai\u0307")', turkish), 'Aİ');
  assert.equal(calc('PROPER("İSTANBUL IŞIK")', turkish), 'İstanbul Işık');
  assert.equal(calc('PROPER("I\u0307STANBUL")', turkish), 'İstanbul', 'a capital with its dot written after it');
  assert.equal(calc('PROPER("i\u0307stanbul")', turkish), 'İstanbul');
  assert.equal(calc('PROPER("kaI\u0307n")', turkish), 'Kain');
});

test('a combining mark does not end a word in PROPER, whatever the locale', () => {
  for (const engine of [blank, withLocale(LOCALE_EN), withLocale(LOCALE_TR)]) {
    const proper = calc('PROPER("İSTANBUL")', engine);
    assert.equal(typeof proper, 'string', 'PROPER gives text');
    assert.equal((proper as string).slice(-7), 'stanbul', `the letters after the dotted capital are small: ${proper}`);
    assert.equal(calc('PROPER("E\u0301COLE NORMALE")', engine), 'E\u0301cole Normale', 'an accent written after its letter');
    assert.equal(String(calc('PROPER("I\u0307STANBUL")', engine)).slice(-7), 'stanbul', 'a dot written after the capital');
  }
});

/** PROPER, UPPER and LOWER as they were before a locale was known. */
function oldProper(text: string): string {
  let out = '';
  let inWord = false;
  for (let i = 0; i < text.length; i++) {
    const c = text.charAt(i);
    out += inWord ? c.toLowerCase() : c.toUpperCase();
    inWord = c.toLowerCase() !== c.toUpperCase();
  }
  return out;
}

test('without a Turkish locale UPPER, LOWER and PROPER give what they always gave for ASCII and Latin-1', () => {
  let seed = 4711;
  const random = () => ((seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0) / 4294967296);
  const engines = [blank, withLocale(LOCALE_EN), withLocale(undefined)];
  for (let round = 0; round < 400; round++) {
    let text = '';
    const length = Math.floor(random() * 24);
    for (let i = 0; i < length; i++) {
      // Printable ASCII and Latin-1, with spaces often enough to make words; no quote, which would end the literal.
      let code = random() < 0.15 ? 32 : random() < 0.6 ? 33 + Math.floor(random() * 94) : 160 + Math.floor(random() * 96);
      if (code === 34) code = 39;
      text += String.fromCharCode(code);
    }
    for (const engine of engines) {
      assert.equal(calc(`UPPER("${text}")`, engine), text.toUpperCase(), `UPPER of ${JSON.stringify(text)}`);
      assert.equal(calc(`LOWER("${text}")`, engine), text.toLowerCase(), `LOWER of ${JSON.stringify(text)}`);
      assert.equal(calc(`PROPER("${text}")`, engine), oldProper(text), `PROPER of ${JSON.stringify(text)}`);
    }
  }
});

test('a name the formula of the cell already calls is not read as a Turkish name', () => {
  // VAL, S, KOD ... are functions of a macro or an add-in in this file; in Turkish they are DGET, N, CODE.
  for (const [stored, shown] of [['VAL(A1)', 'VAL(A1)'], ['S(A1)+M(B1)', 'S(A1)+M(B1)'], ['KOD(TEK(1),"x")', 'KOD(TEK(1);"x")'], ['IF(DA(1,2)>BD(1.5),GD(1),YAT(2))', 'IF(DA(1;2)>BD(1,5);GD(1);YAT(2))'], ['TİP(1)', 'TİP(1)']]) {
    assert.equal(formulaFromLocal(shown, LOCALE_TR, { stored }), stored, `${stored} typed again as it is shown`);
  }
  // What the user adds is read the Turkish way; what was there stays.
  assert.equal(formulaFromLocal('VAL(A1)+TOPLA(B1;2)+KOD("a")', LOCALE_TR, { stored: 'VAL(A1)' }), 'VAL(A1)+SUM(B1,2)+CODE("a")');
  // The name in a text or in the name of a sheet of the stored formula is no call.
  assert.equal(formulaFromLocal('VAL(A1)', LOCALE_TR, { stored: '"VAL("&\'VAL(\'!A1' }), 'DGET(A1)');
  // Without a stored formula, and where the stored one does not call the name, nothing changes.
  assert.equal(formulaFromLocal('VAL(A1)', LOCALE_TR), 'DGET(A1)');
  assert.equal(formulaFromLocal('VAL(A1)', LOCALE_TR, { stored: 'SUM(A1)' }), 'DGET(A1)');
  // A name that is the English one but for its dots is still stored the English way.
  assert.equal(formulaFromLocal('SİN(A1)', LOCALE_TR, { stored: 'SIN(A1)' }), 'SIN(A1)');
  assert.equal(formulaFromLocal('MİN(A1)+MOD(1;2)', LOCALE_TR, { stored: 'MIN(A1)+MOD(1,2)' }), 'MIN(A1)+MOD(1,2)');
});

test('a function is not read as a Turkish name where the workbook has a defined name of that spelling', () => {
  const isName = (name: string) => name.toUpperCase() === 'VAL';
  assert.equal(formulaFromLocal('VAL(A1)+KOD("a")', LOCALE_TR, { isName }), 'VAL(A1)+CODE("a")');
  assert.equal(formulaFromLocal('val(A1)', LOCALE_TR, { isName }), 'val(A1)');
  // Only the short names that a function of a file may well be called. Any other Turkish name is the
  // function, as in Excel, whatever the workbook has defined.
  const any = () => true;
  assert.equal(formulaFromLocal('TOPLA(1;2)+EĞER(A1;1;2)+AY(A1)+MAK(1;2)+ÜS(1)', LOCALE_TR, { isName: any }), 'SUM(1,2)+IF(A1,1,2)+MONTH(A1)+MAX(1,2)+EXP(1)');
  for (const short of ['VAL', 'S', 'M', 'KOD', 'TEK', 'TİP', 'DA', 'BD', 'GD', 'YAT']) {
    assert.equal(formulaFromLocal(`${short}(1)`, LOCALE_TR, { isName: any }), `${short}(1)`, short);
  }
});

test('a formula of the file typed again as it is shown stays the formula it was', () => {
  const workbook = book({ A1: 5, B1: '=VAL(A1)+S(A1)', B2: '=SUM(A1,1.5)', B3: '=DA(A1)' });
  workbook.definedNames.push({ name: 'DA', formula: 'S!$A$1' } as never);
  const sheet = workbook.sheets[0];
  for (const row of [0, 1, 2]) {
    const before = getCell(sheet, row, 1)!.formula!.text;
    const shown = '=' + formulaToLocalText(before);
    const ops = setCellInput(workbook, 0, row, 1, shown, { locale: LOCALE_TR });
    const cell = cellOf(ops);
    assert.equal(cell, before, `${shown} typed again`);
  }
  // A cell without a formula: the name is the Turkish one.
  assert.equal(cellOf(setCellInput(workbook, 0, 5, 1, '=VAL(A1)', { locale: LOCALE_TR })), 'DGET(A1)');
  // The defined name DA keeps a call of that spelling from being read as SLN, in any cell.
  assert.equal(cellOf(setCellInput(workbook, 0, 6, 1, '=DA(1;2;3)', { locale: LOCALE_TR })), 'DA(1,2,3)');
});

test('the engine of a view asks the view for its locale, each time', () => {
  let locale: FormatLocale = LOCALE_TR;
  const options = viewEngineOptions(() => locale);
  assert.deepEqual(options.budget, { milliseconds: 250 });
  const workbook = book({ A1: '=UPPER("i")' });
  const engine = new CalcEngine(workbook, { ...options, budget: { milliseconds: 0 } });
  engine.recalculate(undefined, { overwrite: true });
  assert.equal(getCell(workbook.sheets[0], 0, 0)!.value, 'İ');
  locale = LOCALE_EN;
  engine.recalculate(undefined, { overwrite: true });
  assert.equal(getCell(workbook.sheets[0], 0, 0)!.value, 'I');
  assert.equal(options.locale!(), LOCALE_EN);
});
