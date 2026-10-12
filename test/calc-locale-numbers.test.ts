// Tests for numbers written the way the locale of the view writes them, inside the text a
// formula works with: criteria (">7,5"), VALUE, NUMBERVALUE, text met by an operator.
//
// The rule that must hold: text that is a number or a date the way files write them has the
// same value in every locale. A locale with a decimal comma only adds readings for text
// that was no number ("7,5", "1.234,5"). What functions write (TEXT, FIXED, DOLLAR, "&")
// does not depend on the locale, so the engine always reads back what it wrote.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine, type CalcEngineOptions } from '../src/calc/engine';
import { numberToText, textToNumber, type CalcContext } from '../src/calc/values';
import { parseCellInput } from '../src/edit/input';
import { LOCALE_EN, LOCALE_TR, type FormatLocale } from '../src/format/locale';
import { serialToDate } from '../src/format/number-format';
import { setCellInput } from '../src/ops/commands';
import { applyOp } from '../src/ops/ops';
import { createSheet, emptyStyleSheet, getCell, parseCellRef, type Cell, type CellFormula, type Row, type Workbook } from '../src/model/sheet/types';
import { parseNumberText } from '../src/calc/functions/helpers';
import * as tables from './calc-locale-numbers-tables';

type Plain = number | string | boolean | null;
interface Stored {
  f: string;
  v: Plain;
}
type Cells = Record<string, Plain | Stored>;

function makeCell(spec: Plain | Stored): Cell {
  if (typeof spec === 'string' && spec.charAt(0) === '=') {
    const formula: CellFormula = { text: spec.slice(1), type: 'normal' };
    return { kind: 'blank', value: null, style: 0, formula };
  }
  if (spec !== null && typeof spec === 'object') {
    const formula: CellFormula = { text: spec.f, type: 'normal' };
    const kind = spec.v === null ? 'blank' : typeof spec.v === 'number' ? 'number' : typeof spec.v === 'boolean' ? 'boolean' : 'string';
    return { kind, value: spec.v, style: 0, formula };
  }
  return { kind: spec === null ? 'blank' : typeof spec === 'number' ? 'number' : typeof spec === 'boolean' ? 'boolean' : 'string', value: spec, style: 0 };
}

function putCell(workbook: Workbook, address: string, spec: Plain | Stored): { sheet: number; row: number; col: number } {
  const sheet = workbook.sheets[0];
  const at = parseCellRef(address) as { row: number; col: number };
  // A new row object, as an edit makes one.
  const old = sheet.rows.get(at.row);
  const row = { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: old ? old.cells.slice() : [] } as Row;
  row.cells[at.col] = makeCell(spec);
  sheet.rows.set(at.row, row);
  if (at.row >= sheet.rowCount) sheet.rowCount = at.row + 1;
  if (at.col >= sheet.colCount) sheet.colCount = at.col + 1;
  return { sheet: 0, row: at.row, col: at.col };
}

function book(cells: Cells = {}): Workbook {
  const workbook: Workbook = { sheets: [], styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 };
  workbook.sheets.push(createSheet({ name: 'S', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' }));
  for (const address of Object.keys(cells)) putCell(workbook, address, cells[address]);
  return workbook;
}

function valueOf(workbook: Workbook, address: string): unknown {
  const at = parseCellRef(address) as { row: number; col: number };
  const cell = getCell(workbook.sheets[0], at.row, at.col);
  return cell ? cell.value : undefined;
}

/** The clock the tables of results were written down with. */
const NO_LIMIT: CalcEngineOptions = { budget: { milliseconds: 0 }, now: () => 45356.5 };
const AWAY = { sheet: 0, row: 60, col: 20 };
const VALUE_ERROR = { error: '#VALUE!' };

function engineOf(locale: FormatLocale | undefined | 'none', cells: Cells = {}): CalcEngine {
  return new CalcEngine(book(cells), locale === 'none' ? NO_LIMIT : { ...NO_LIMIT, locale: () => locale });
}

function calc(formula: string, engine: CalcEngine): unknown {
  return engine.evaluateText(formula, AWAY);
}

/** The three engines that read numbers the way files write them: no locale at all, a locale that is not there, English. */
function plainEngines(cells: Cells = {}): [string, CalcEngine][] {
  return [
    ['no locale', engineOf('none', cells)],
    ['undefined locale', engineOf(undefined, cells)],
    ['English', engineOf(LOCALE_EN, cells)],
  ];
}

// A1:A6 numbers, B1:B6 what is summed, C1:C6 a second column of conditions.
const DATA: Cells = {
  A1: 7, A2: 7.5, A3: 8, A4: 1234.5, A5: 7.25, A6: 'x7,5',
  B1: 1, B2: 10, B3: 100, B4: 1000, B5: 10000, B6: 100000,
  C1: 'a', C2: 'b', C3: 'a', C4: 'a', C5: 'b', C6: 'a',
};


// ---------------------------------------------------------------------------
// Criteria

test('in a Turkish locale the number of a criterion may be written with a decimal comma', () => {
  const tr = engineOf(LOCALE_TR, DATA);
  assert.equal(calc('COUNTIF(A1:A6,">7,5")', tr), 2, '">7,5": 8 and 1234.5');
  assert.equal(calc('COUNTIF(A1:A6,">=7,5")', tr), 3, '">=7,5"');
  assert.equal(calc('COUNTIF(A1:A6,"<7,5")', tr), 2, '"<7,5": 7 and 7.25');
  assert.equal(calc('COUNTIF(A1:A6,"=7,5")', tr), 1, '"=7,5"');
  assert.equal(calc('COUNTIF(A1:A6,"7,5")', tr), 1, '"7,5" without an operator asks for equality');
  assert.equal(calc('COUNTIF(A1:A6,"<>7,5")', tr), 5, '"<>7,5": everything but the 7.5');
  assert.equal(calc('COUNTIF(A1:A6,"<=1.234,5")', tr), 5, '"<=1.234,5": thousands separator and decimal comma');
  assert.equal(calc('COUNTIF(A1:A6,">7,2%")', tr), 5, '">7,2%" is 0.072');
});

test('a criterion that is a number the way files write it means the same in every locale', () => {
  for (const [name, engine] of plainEngines(DATA).concat([['Turkish', engineOf(LOCALE_TR, DATA)]])) {
    assert.equal(calc('COUNTIF(A1:A6,">7.5")', engine), 2, `${name}: ">7.5"`);
    assert.equal(calc('COUNTIF(A1:A6,"=7.25")', engine), 1, `${name}: "=7.25"`);
    assert.equal(calc('COUNTIF(A1:A6,">1,234.4")', engine), 1, `${name}: ">1,234.4"`);
    assert.equal(calc('COUNTIF(A1:A6,">7")', engine), 4, `${name}: ">7"`);
    // "1.234" is one and a bit, "1,234" a thousand and more: as a file means them.
    assert.equal(calc('COUNTIF(A1:A6,">1.234")', engine), 5, `${name}: ">1.234"`);
    assert.equal(calc('COUNTIF(A1:A6,">1,234")', engine), 1, `${name}: ">1,234"`);
    assert.equal(calc('COUNTIF(A1:A6,">"&1.234)', engine), 5, `${name}: a number joined to the operator`);
    assert.equal(calc('COUNTIF(A1:A6,">"&7.5)', engine), 2, `${name}: ">"&7.5`);
  }
});

test('without a Turkish locale a criterion with a comma is read as before', () => {
  for (const [name, engine] of plainEngines(DATA)) {
    // "7,5" is the two digits 75 with a thousands separator out of place: no number, so text.
    assert.equal(calc('COUNTIF(A1:A6,">7,5")', engine), 1, `${name}: ">7,5" compares text`);
    assert.equal(calc('COUNTIF(A1:A6,"=7,5")', engine), 0, `${name}: "=7,5"`);
    assert.equal(calc('COUNTIF(A1:A6,"<>7,5")', engine), 6, `${name}: "<>7,5"`);
  }
});

test('a text criterion with a wildcard does not depend on the locale', () => {
  const cells: Cells = { A1: 'x7,5', A2: 'x7.5', A3: 'a,b', A4: '7,5 kg', A5: 7.5 };
  for (const [name, engine] of plainEngines(cells).concat([['Turkish', engineOf(LOCALE_TR, cells)]])) {
    assert.equal(calc('COUNTIF(A1:A5,"x7,*")', engine), 1, `${name}: "x7,*"`);
    assert.equal(calc('COUNTIF(A1:A5,"*,*")', engine), 3, `${name}: "*,*"`);
    assert.equal(calc('COUNTIF(A1:A5,"?7?5")', engine), 2, `${name}: "?7?5"`);
    assert.equal(calc('COUNTIF(A1:A5,"a,b")', engine), 1, `${name}: "a,b"`);
    assert.equal(calc('COUNTIF(A1:A5,"7,5 kg")', engine), 1, `${name}: "7,5 kg"`);
  }
});

test('in a Turkish locale an "equals" criterion also takes text cells that read as the number', () => {
  const cells: Cells = { A1: 7.5, A2: '7,5', A3: '7.5', A4: '7,50', A5: 'x', A6: '1.234', A7: 1234, A8: '1,234' };
  const tr = engineOf(LOCALE_TR, cells);
  assert.equal(calc('COUNTIF(A1:A8,"7,5")', tr), 4, 'the number, "7,5", "7.5" and "7,50"');
  assert.equal(calc('COUNTIF(A1:A8,7.5)', tr), 4, 'a number as the criterion');
  assert.equal(calc('COUNTIF(A1:A8,"<>7,5")', tr), 4, 'the rest');
  // What a file means by them: "1.234" is 1.234, and "1,234" is no plain number at all.
  assert.equal(calc('COUNTIF(A1:A8,"1.234")', tr), 1, 'text "1.234" is 1.234');
  assert.equal(calc('COUNTIF(A1:A8,1234)', tr), 1, 'only the number 1234');
  // Asked more than once in one recalculation, the answer comes from the index of the range: the same one.
  const workbook = book({ ...cells, C1: '=COUNTIF(A1:A8,"7,5")', C2: '=COUNTIF(A1:A8,"7,5")', C3: '=COUNTIF(A1:A8,7.5)', C4: '=COUNTIF(A1:A8,1234)', C5: '=COUNTIF(A1:A8,1234)' });
  new CalcEngine(workbook, { ...NO_LIMIT, locale: () => LOCALE_TR }).recalculate(undefined, { overwrite: true });
  assert.deepEqual(['C1', 'C2', 'C3', 'C4', 'C5'].map((cell) => valueOf(workbook, cell)), [4, 4, 4, 1, 1]);
  for (const [name, engine] of plainEngines(cells)) {
    assert.equal(calc('COUNTIF(A1:A8,7.5)', engine), 2, `${name}: the number and "7.5"`);
    assert.equal(calc('COUNTIF(A1:A8,"7,5")', engine), 1, `${name}: the text "7,5"`);
    assert.equal(calc('COUNTIF(A1:A8,1234)', engine), 1, `${name}: the number 1234`);
  }
});

test('the family of criteria functions reads a criterion the same way', () => {
  const tr = engineOf(LOCALE_TR, DATA);
  assert.equal(calc('SUMIF(A1:A6,">7,5",B1:B6)', tr), 1100, 'SUMIF');
  assert.equal(calc('SUMIFS(B1:B6,A1:A6,">7,2",C1:C6,"a")', tr), 1100, 'SUMIFS');
  assert.equal(calc('COUNTIFS(A1:A6,">=7,25",A1:A6,"<=1.234,5")', tr), 4, 'COUNTIFS');
  assert.equal(calc('AVERAGEIF(A1:A6,"<7,5")', tr), 7.125, 'AVERAGEIF');
  assert.equal(calc('AVERAGEIFS(B1:B6,A1:A6,"<=7,5",C1:C6,"b")', tr), 5005, 'AVERAGEIFS');
  assert.equal(calc('MAXIFS(A1:A6,A1:A6,"<7,5")', tr), 7.25, 'MAXIFS');
  assert.equal(calc('MINIFS(A1:A6,A1:A6,">7,5")', tr), 8, 'MINIFS');
  for (const [name, engine] of plainEngines(DATA)) {
    assert.equal(calc('SUMIF(A1:A6,">7.5",B1:B6)', engine), 1100, `${name}: SUMIF`);
    assert.equal(calc('SUMIF(A1:A6,">7,5",B1:B6)', engine), 100000, `${name}: SUMIF with ">7,5" compares text`);
    assert.equal(calc('SUMIFS(B1:B6,A1:A6,">7,2",C1:C6,"a")', engine), 100000, `${name}: SUMIFS`);
    assert.equal(calc('MAXIFS(A1:A6,A1:A6,"<7.5")', engine), 7.25, `${name}: MAXIFS`);
    assert.equal(calc('MAXIFS(A1:A6,A1:A6,"<7,5")', engine), 0, `${name}: MAXIFS with "<7,5"`);
  }
});

test('the criteria cells of the database functions read numbers the way the locale writes them', () => {
  const cells: Cells = {
    A1: 'Size', B1: 'Sum',
    A2: 7, B2: 1,
    A3: 7.5, B3: 10,
    A4: 8, B4: 100,
    A5: 1234.5, B5: 1000,
    D1: 'Size', D2: '>7,5',
    E1: 'Size', E2: '7,5',
    F1: 'Size', F2: '<=1.234,5', G1: 'Size', G2: '>7.4',
  };
  const tr = engineOf(LOCALE_TR, cells);
  assert.equal(calc('DSUM(A1:B5,"Sum",D1:D2)', tr), 1100, 'DSUM with ">7,5"');
  assert.equal(calc('DSUM(A1:B5,"Sum",E1:E2)', tr), 10, 'DSUM with "7,5": the number, not text that begins so');
  assert.equal(calc('DSUM(A1:B5,"Sum",F1:G2)', tr), 1110, 'DSUM with two conditions in one row');
  assert.equal(calc('DCOUNT(A1:B5,"Sum",D1:D2)', tr), 2, 'DCOUNT');
  for (const [name, engine] of plainEngines(cells)) {
    assert.equal(calc('DSUM(A1:B5,"Sum",D1:D2)', engine), 0, `${name}: ">7,5" compares text`);
    assert.equal(calc('DSUM(A1:B5,"Sum",G1:G2)', engine), 1110, `${name}: ">7.4"`);
  }
});

test('criteria typed in a Turkish view: the comma inside the string, and a number joined to the operator', () => {
  const workbook = book({ A1: 7, A2: 7.5, A3: 8, A4: 9.25 });
  const type = (col: number, text: string) => {
    for (const op of setCellInput(workbook, 0, 0, col, text, { locale: LOCALE_TR })) applyOp(workbook, op);
    const cell = getCell(workbook.sheets[0], 0, col) as Cell;
    return cell.formula && cell.formula.text;
  };
  assert.equal(type(2, '=EĞERSAY(A1:A4;">7,5")'), 'COUNTIF(A1:A4,">7,5")', 'what is inside the string is not rewritten');
  assert.equal(type(3, '=EĞERSAY(A1:A4;">"&7,5)'), 'COUNTIF(A1:A4,">"&7.5)', 'the number outside the string is');
  assert.equal(type(4, '=EĞERSAY(A:A;">"&7,5)'), 'COUNTIF(A:A,">"&7.5)', 'a whole column');
  new CalcEngine(workbook, { ...NO_LIMIT, locale: () => LOCALE_TR }).recalculate(undefined, { overwrite: true });
  assert.deepEqual([valueOf(workbook, 'C1'), valueOf(workbook, 'D1'), valueOf(workbook, 'E1')], [2, 2, 2]);
});

test('what the engine joins to an operator is read back as the number it was, in a Turkish locale too', () => {
  const cells: Cells = { A1: 0.125, A2: 1.234, A3: 7.5, A4: 125, A5: 1234, A6: 0.5, A7: 0.4, A8: 0.1, B1: 1.234, B2: 0.125 };
  for (const [name, engine] of plainEngines(cells).concat([['Turkish', engineOf(LOCALE_TR, cells)]])) {
    assert.equal(calc('COUNTIF(A1:A5,">"&B1)', engine), 3, `${name}: ">"&B1 with 1.234 in B1`);
    assert.equal(calc('COUNTIF(A1:A5,"="&B2)', engine), 1, `${name}: "="&B2 with 0.125 in B2`);
    assert.equal(calc('COUNTIF(A1:A5,"<="&B1)', engine), 2, `${name}: "<="&B1`);
    assert.equal(calc('SUMIF(A1:A5,">"&B2)', engine), 1367.734, `${name}: SUMIF`);
    assert.equal(calc('COUNTIF(A1:A8,">"&TIME(3,0,0))', engine), 6, `${name}: a time joined to the operator`);
    assert.equal(calc('(0.125&"")+0', engine), 0.125, `${name}: (0.125&"")+0`);
    assert.equal(calc('VALUE(1.234&"")', engine), 1.234, `${name}: VALUE(1.234&"")`);
    assert.equal(calc('VALUE(LEFT(B1,5))', engine), 1.234, `${name}: VALUE(LEFT(B1,5))`);
    assert.equal(calc('VALUE(TEXTJOIN("",TRUE,B2))', engine), 0.125, `${name}: TEXTJOIN`);
  }
});

// ---------------------------------------------------------------------------
// Text as a number

test('VALUE and text met by an operator read the decimal comma in a Turkish locale', () => {
  const tr = engineOf(LOCALE_TR, { A1: '1,5', A2: '2.500,25' });
  assert.equal(calc('VALUE("1,5")', tr), 1.5, 'VALUE("1,5")');
  assert.equal(calc('VALUE("0,5")', tr), 0.5, 'VALUE("0,5")');
  assert.equal(calc('VALUE("1.234,5")', tr), 1234.5, 'VALUE("1.234,5")');
  assert.equal(calc('VALUE("1.234.567,89")', tr), 1234567.89, 'two thousands separators');
  assert.equal(calc('VALUE("-1,5")', tr), -1.5, 'a sign');
  assert.equal(calc('VALUE(" 12,5% ")', tr), 0.125, 'a percentage');
  assert.equal(calc('VALUE("1,5E+3")', tr), 1500, 'an exponent');
  assert.equal(calc('VALUE(",5")', tr), 0.5, 'no digit before the comma');
  assert.equal(calc('VALUE("$1.234,5")', tr), 1234.5, 'a currency sign');
  assert.equal(calc('VALUE("(1,5)")', tr), -1.5, 'parentheses');
  assert.equal(calc('"1,5"+1', tr), 2.5, 'text met by "+"');
  assert.equal(calc('A1*2', tr), 3, 'a text cell met by "*"');
  assert.equal(calc('A2-0.25', tr), 2500, 'a text cell with both separators');
  assert.equal(calc('-"1,5"', tr), -1.5, 'a sign in front of text');
  assert.equal(calc('"50,5%"*2', tr), 1.01, 'a percentage as text');
  assert.equal(calc('SUM("1,5",2)', tr), 3.5, 'text written into SUM');
  assert.equal(calc('COUNT("1,5","x")', tr), 1, 'COUNT of text written into it');
  assert.equal(calc('ROUND("2,44",1)', tr), 2.4, 'a numeric argument');
  assert.equal(calc('ROUND("2,449",1)', tr), 2449, 'three digits after a comma are thousands, as a file means them');
  assert.equal(calc('FIXED("1.234,567",2)', tr), '1,234.57', 'a numeric argument of FIXED');
});

test('text that is a number or a date the way files write it means the same in a Turkish locale', () => {
  const tr = engineOf(LOCALE_TR);
  const none = engineOf('none');
  assert.equal(calc('VALUE("1.5")', tr), 1.5, '"1.5"');
  assert.equal(calc('VALUE("1.234")', tr), 1.234, '"1.234" is not a thousand');
  assert.equal(calc('VALUE("0.125")', tr), 0.125, '"0.125"');
  assert.equal(calc('VALUE("7.500")', tr), 7.5, '"7.500"');
  assert.equal(calc('VALUE("1,234")', tr), 1234, '"1,234" keeps its thousands separator');
  assert.equal(calc('VALUE("12,345")', tr), 12345, '"12,345"');
  assert.equal(calc('"7.25"+1', tr), 8.25, '"7.25"+1');
  assert.equal(calc('VALUE("1,234.5")', tr), 1234.5, 'English separators');
  assert.equal(calc('VALUE("1.5E3")', tr), 1500, 'an exponent after a point');
  assert.equal(calc('VALUE("1 1/2")', tr), 1.5, 'a fraction');
  assert.equal(calc('VALUE("12:00")', tr), 0.5, 'a time');
  for (const text of ['3/4/2024', 'March 5, 2024', 'Mar 5,2024', '2024-03-05', '12:30:15', '12:30.5', '1.2.2024', '15.03.2024', '1/2', '10,30 AM', '3,4/2024', '12:30,5', '1,2:30']) {
    assert.deepEqual(calc(`VALUE("${text}")`, tr), calc(`VALUE("${text}")`, none), `VALUE of "${text}"`);
    assert.deepEqual(calc(`DATEVALUE("${text}")`, tr), calc(`DATEVALUE("${text}")`, none), `DATEVALUE of "${text}"`);
  }
});

test('in a Turkish locale text the locale would not write as a number stays text', () => {
  const tr = engineOf(LOCALE_TR);
  for (const text of ['1,5,5', 'abc', '1,5x', 'a,b', '5,', '1.234,', '1.234.', ',', '.', '1..234', '1.23,4', '1.2345,6', '1,5.2024', '1.,5', '1,,5', '1,5-', '--1,5', '1,5e', '1,5 1/2', '1 000,5']) {
    assert.deepEqual(calc(`VALUE("${text}")`, tr), VALUE_ERROR, `VALUE("${text}")`);
  }
  assert.deepEqual(calc('"a,b"+1', tr), VALUE_ERROR, 'text with a comma met by an operator');
});

test('without a Turkish locale VALUE and operators read text as before', () => {
  for (const [name, engine] of plainEngines()) {
    assert.equal(calc('VALUE("1.5")', engine), 1.5, `${name}: "1.5"`);
    assert.deepEqual(calc('VALUE("1,5")', engine), VALUE_ERROR, `${name}: "1,5" is no number`);
    assert.equal(calc('VALUE("1,234")', engine), 1234, `${name}: "1,234"`);
    assert.equal(calc('VALUE("1.234")', engine), 1.234, `${name}: "1.234"`);
    assert.deepEqual(calc('"1,5"+1', engine), VALUE_ERROR, `${name}: "1,5"+1`);
    assert.deepEqual(calc('VALUE("1.234,5")', engine), VALUE_ERROR, `${name}: "1.234,5"`);
  }
});

test('NUMBERVALUE without separators reads the way files write first, then the way the locale does; given separators win', () => {
  const tr = engineOf(LOCALE_TR);
  assert.equal(calc('NUMBERVALUE("1.5")', tr), 1.5, 'a number the way files write it');
  assert.equal(calc('NUMBERVALUE("0.125")', tr), 0.125, '"0.125"');
  assert.equal(calc('NUMBERVALUE("1,5")', tr), 15, 'a comma is a thousands separator to NUMBERVALUE, wherever it stands: as before');
  assert.equal(calc('NUMBERVALUE("1,234.5")', tr), 1234.5, 'English separators');
  assert.equal(calc('NUMBERVALUE("1.234,5")', tr), 1234.5, 'no number the way files write it: the separators of the locale');
  assert.equal(calc('NUMBERVALUE("1.234,5%")', tr), 12.345, 'the same with a percent sign');
  // The locale is asked only when its decimal sign is in the text: points alone stay what they were.
  assert.deepEqual(calc('NUMBERVALUE("1.234.5")', tr), VALUE_ERROR, 'two points and no comma: an error, as before');
  assert.deepEqual(calc('NUMBERVALUE("1.2.3")', tr), VALUE_ERROR, 'points alone');
  assert.deepEqual(calc('NUMBERVALUE("1.000.000")', tr), VALUE_ERROR, 'thousands points without a decimal comma');
  assert.equal(calc('NUMBERVALUE("1.000.000,5")', tr), 1000000.5, 'with the decimal comma they are thousands');
  assert.deepEqual(calc('NUMBERVALUE("1,5,5")', tr), 155, 'commas before the point are dropped, as before');
  assert.deepEqual(calc('NUMBERVALUE("abc")', tr), VALUE_ERROR, 'text');
  // Given separators: exactly what the function always did, in every locale.
  for (const [name, engine] of plainEngines().concat([['Turkish', tr]])) {
    assert.equal(calc('NUMBERVALUE("1,234.5",".",",")', engine), 1234.5, `${name}: both given`);
    assert.equal(calc('NUMBERVALUE("1.234,5",",",".")', engine), 1234.5, `${name}: both given, the Turkish way`);
    assert.deepEqual(calc('NUMBERVALUE("1.234,5",".",",")', engine), VALUE_ERROR, `${name}: given separators the text does not follow`);
    assert.equal(calc('NUMBERVALUE("1.5",".")', engine), 1.5, `${name}: a given decimal point`);
    assert.deepEqual(calc('NUMBERVALUE("1,5",",")', engine), VALUE_ERROR, `${name}: a given decimal comma meets the thousands comma that was not given`);
    assert.deepEqual(calc('NUMBERVALUE("1.5",,".")', engine), VALUE_ERROR, `${name}: a given thousands point meets the decimal point that was not given`);
    assert.equal(calc('NUMBERVALUE("1 234,5",","," ")', engine), 1234.5, `${name}: a given thousands separator`);
    assert.deepEqual(calc('NUMBERVALUE("1,5",",",",")', engine), VALUE_ERROR, `${name}: the same sign for both`);
  }
  for (const [name, engine] of plainEngines()) {
    assert.equal(calc('NUMBERVALUE("1.5")', engine), 1.5, `${name}: "1.5"`);
    assert.equal(calc('NUMBERVALUE("1,5")', engine), 15, `${name}: "1,5"`);
    assert.deepEqual(calc('NUMBERVALUE("1.234,5")', engine), VALUE_ERROR, `${name}: "1.234,5"`);
  }
});

// ---------------------------------------------------------------------------
// Numbers as text

test('what TEXT, FIXED, DOLLAR and "&" write does not depend on the locale', () => {
  const formulas = [
    ['TEXT(1234.5,"#,##0.00")', '1,234.50'],
    ['TEXT(0.125,"0.0%")', '12.5%'],
    ['TEXT(1234.5,"0.00")', '1234.50'],
    ['TEXT(-1234.5,"#,##0.0;(#,##0.0)")', '(1,234.5)'],
    ['TEXT(1.5,"General")', '1.5'],
    // 5 March 2024 is a Tuesday.
    ['TEXT(45356,"d mmmm yyyy")', '5 March 2024'],
    ['TEXT(45356,"dddd")', 'Tuesday'],
    ['TEXT(45356,"ddd, mmm d")', 'Tue, Mar 5'],
    ['TEXT(45356,"mm/yyyy")', '03/2024'],
    ['TEXT(45356,"m/d")', '3/5'],
    ['TEXT(45356.75,"h:mm AM/PM")', '6:00 PM'],
    ['TEXT(45356,"[$-41F]mmmm")', 'Mart'],
    ['TEXT(45356,"[$-409]mmmm")', 'March'],
    ['TEXT(TRUE,"0.0")', 'TRUE'],
    ['FIXED(1234.567)', '1,234.57'],
    ['FIXED(1234.567,1,TRUE)', '1234.6'],
    ['FIXED(-1234567.891,0)', '-1,234,568'],
    ['DOLLAR(1234.567)', '$1,234.57'],
    ['DOLLAR(-1234.567,1)', '($1,234.6)'],
    ['1.5&""', '1.5'],
    ['CONCAT(0.125,"x")', '0.125x'],
  ];
  for (const [name, engine] of plainEngines().concat([['Turkish', engineOf(LOCALE_TR)]])) {
    for (const [formula, expected] of formulas) assert.equal(calc(formula, engine), expected, `${name}: ${formula}`);
  }
});

test('in a Turkish locale what TEXT, FIXED and DOLLAR write is read back as the same number or date', () => {
  const tr = engineOf(LOCALE_TR);
  const en = engineOf(LOCALE_EN);
  assert.equal(calc('VALUE(TEXT(1234.5,"#,##0.00"))', tr), 1234.5, 'thousands separator and decimals');
  assert.equal(calc('VALUE(TEXT(1.234,"0.000"))', tr), 1.234, 'three decimals are not thousands');
  assert.equal(calc('VALUE(TEXT(0.125,"0.000"))', tr), 0.125, '0.125');
  assert.equal(calc('VALUE(TEXT(45356,"m/d"))', tr), 45356, 'month and day');
  assert.equal(calc('VALUE(TEXT(45356,"mm/yyyy"))', tr), calc('VALUE(TEXT(45356,"mm/yyyy"))', en), 'month and year');
  assert.equal(calc('DATEVALUE(TEXT(45356,"d mmmm yyyy"))', tr), 45356, 'the name of a month');
  assert.equal(calc('DATEVALUE(TEXT(45356,"mm/dd/yyyy"))', tr), 45356, 'a date in digits');
  assert.equal(calc('VALUE(TEXT(0.75,"h:mm AM/PM"))', tr), 0.75, 'a time with PM');
  assert.equal(calc('VALUE(FIXED(1234.567))', tr), 1234.57, 'FIXED');
  assert.equal(calc('VALUE(FIXED(1.234,3))', tr), 1.234, 'FIXED with three decimals');
  assert.equal(calc('VALUE(DOLLAR(-1234.567))', tr), -1234.57, 'DOLLAR');

  const random = generator(4242);
  const codes = ['0', '0.0', '0.000', '#,##0', '#,##0.00', '#,##0.000', '0.00%', '0.0E+00', 'General', '#,##0.0;(#,##0.0)', '$#,##0.00', 'm/d', 'mm/yyyy', 'm/d/yyyy', 'd mmmm yyyy', 'dd-mmm-yy', 'yyyy-mm-dd', 'dd.mm.yyyy', 'h:mm AM/PM', 'hh:mm:ss', 'm/d/yyyy h:mm', 'mmm d, yyyy'];
  for (let i = 0; i < 1500; i++) {
    const kind = Math.floor(random() * 5);
    const x = kind === 0 ? Math.round(random() * 2e6) / 1000 : kind === 1 ? Math.round(random() * 1000) / 1000 : kind === 2 ? 40000 + Math.floor(random() * 8000) + Math.round(random() * 1440) / 1440 : kind === 3 ? -Math.round(random() * 1e7) / 1000 : Math.round(random() * 1e9);
    const code = codes[Math.floor(random() * codes.length)];
    for (const formula of [`TEXT(${x},"${code}")`, `IFERROR(VALUE(TEXT(${x},"${code}")),"e")`, `IFERROR(DATEVALUE(TEXT(${x},"${code}")),"e")`, `IFERROR(TEXT(${x},"${code}")+0,"e")`, `IFERROR(VALUE(FIXED(${x},${i % 5})),"e")`, `IFERROR(VALUE(DOLLAR(${x},${i % 4})),"e")`]) {
      assert.deepEqual(calc(formula, tr), calc(formula, en), formula);
    }
  }
});

// ---------------------------------------------------------------------------
// What the file says stands

test('a value of the file stands in a Turkish view until something the formula reads changes', () => {
  // As a file written elsewhere would have them: results that differ from what the locale now gives.
  const workbook = book({
    A1: 7, A2: 7.5, A3: 8,
    B1: { f: 'COUNTIF(A1:A3,">7,5")', v: 0 },
    B2: { f: 'VALUE("1,5")+A1', v: 99 },
    // And those whose value is what every locale gives: nothing to report.
    B3: { f: 'TEXT(A2,"0.00")', v: '7.50' },
    B4: { f: 'FIXED(A3*1000,1)', v: '8,000.0' },
    B5: { f: 'COUNTIF(A1:A3,">=7.5")', v: 2 },
  });
  const engine = new CalcEngine(workbook, { ...NO_LIMIT, locale: () => LOCALE_TR });
  const values = () => ['B1', 'B2', 'B3', 'B4', 'B5'].map((cell) => valueOf(workbook, cell));
  engine.recalculate();
  assert.deepEqual(values(), [0, 99, '7.50', '8,000.0', 2], 'the values of the file');
  assert.deepEqual(engine.mismatched.map((cell) => cell.row).sort(), [0, 1], 'the two that differ are reported');

  // A1 changes: what reads it follows the locale.
  engine.recalculate([putCell(workbook, 'A1', 9)]);
  assert.deepEqual(values(), [2, 10.5, '7.50', '8,000.0', 3]);

  engine.recalculate(undefined, { overwrite: true });
  assert.deepEqual(values(), [2, 10.5, '7.50', '8,000.0', 3]);
  assert.deepEqual(engine.mismatched, []);
});

test('in a Turkish view the values a file has for volatile formulas and their like stand', () => {
  // Every formula reads only what a file writes with a point; what the file says is what is calculated.
  const cells: Cells = {
    A1: 0.5, A2: 1.5, A3: 2000, D1: 1.234, Z1: 1,
    B1: { f: 'COUNTIF(A1:A3,">1.234")', v: 2 },
    B2: { f: 'VALUE("1.234")', v: 1.234 },
    B3: { f: 'TEXT(A2,"0.00")', v: '1.50' },
    B4: { f: 'COUNTIF(A1:A3,">"&D1)', v: 2 },
    B5: { f: 'COUNTIF(OFFSET(A1,0,0,3,1),">1.234")', v: 2 },
    B6: { f: 'VALUE("1.234")+0*NOW()', v: 1.234 },
    B7: { f: 'TEXT(TODAY()*0+1.5,"0.00")', v: '1.50' },
    B8: { f: 'VALUE(INDIRECT("D1")&"")', v: 1.234 },
    B9: { f: 'IF(RAND()>2,0,VALUE("1.234"))', v: 1.234 },
  };
  const keys = ['B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'B9'];
  const file = [2, 1.234, '1.50', 2, 2, 1.234, '1.50', 1.234, 1.234];
  for (const locale of [LOCALE_TR, LOCALE_EN]) {
    const name = locale === LOCALE_TR ? 'Turkish' : 'English';
    const workbook = book(cells);
    const engine = new CalcEngine(workbook, { ...NO_LIMIT, locale: () => locale });
    const values = () => keys.map((cell) => valueOf(workbook, cell));
    engine.recalculate([]);
    assert.deepEqual(values(), file, `${name}: a recalculation of nothing (the volatile ones)`);
    engine.recalculate();
    assert.deepEqual(values(), file, `${name}: a recalculation of everything`);
    engine.recalculate([putCell(workbook, 'Z1', 2)]);
    assert.deepEqual(values(), file, `${name}: after an edit nothing reads`);
    engine.recalculate([putCell(workbook, 'A1', 0.6)]);
    assert.deepEqual(values(), file, `${name}: after an edit that changes no result`);
    assert.deepEqual(engine.mismatched, [], `${name}: nothing differs from the file`);
    engine.recalculate(undefined, { overwrite: true });
    assert.deepEqual(values(), file, `${name}: a calculation that overwrites`);
  }
  // Opened with one formula the file has no value for.
  const workbook = book({ ...cells, C1: '=1+1' });
  const engine = new CalcEngine(workbook, { ...NO_LIMIT, locale: () => LOCALE_TR });
  engine.recalculate(engine.uncalculated());
  assert.deepEqual(keys.map((cell) => valueOf(workbook, cell)), file, 'opened with an uncalculated cell');
  assert.equal(valueOf(workbook, 'C1'), 2);
});

test('a change of the locale is followed by the next calculation, not before', () => {
  let locale: FormatLocale = LOCALE_EN;
  const workbook = book({ A1: 7, A2: 7.5, A3: 8, B1: '=COUNTIF(A1:A3,">7,5")', B2: '=TEXT(A2,"0.0")', B3: '=VALUE("1,5")', B4: '=FIXED(A2*1000,1)' });
  const engine = new CalcEngine(workbook, { ...NO_LIMIT, locale: () => locale });
  const values = () => ['B1', 'B2', 'B3', 'B4'].map((cell) => valueOf(workbook, cell));
  engine.recalculate(undefined, { overwrite: true });
  assert.deepEqual(values(), [0, '7.5', '#VALUE!', '7,500.0']);
  locale = LOCALE_TR;
  assert.deepEqual(values(), [0, '7.5', '#VALUE!', '7,500.0'], 'nothing is calculated by the change itself');
  engine.recalculate(undefined, { overwrite: true });
  assert.deepEqual(values(), [1, '7.5', 1.5, '7,500.0']);
  locale = LOCALE_EN;
  engine.recalculate(undefined, { overwrite: true });
  assert.deepEqual(values(), [0, '7.5', '#VALUE!', '7,500.0']);
});

// ---------------------------------------------------------------------------
// Results written down before the locale was read

const TABLES: [string, tables.Cells, tables.Row[], tables.Row[]][] = [
  ['numbers the engine writes, criteria, lookups', tables.T2_CELLS, tables.T2_SAME, tables.T2_GAINED],
  ['TEXT, FIXED and DOLLAR', tables.T6_CELLS, tables.T6_SAME, tables.T6_GAINED],
  ['dates through TEXT, times in criteria, NUMBERVALUE', tables.T8_CELLS, tables.T8_SAME, tables.T8_GAINED],
  ['VALUE and DATEVALUE of text', tables.T1_CELLS, tables.T1_SAME, tables.T1_GAINED],
];

for (const [name, cells, same, gained] of TABLES) {
  test(`in a Turkish locale formulas give what they gave before the locale was read: ${name}`, () => {
    const tr = engineOf(LOCALE_TR, cells);
    const none = engineOf('none', cells);
    const en = engineOf(LOCALE_EN, cells);
    assert.equal(same.length > 20, true, `the table holds rows (${same.length})`);
    for (const [formula, expected] of same) {
      assert.deepEqual(calc(formula, tr), expected, `Turkish: ${formula}`);
      assert.deepEqual(calc(formula, none), expected, `no locale: ${formula}`);
      assert.deepEqual(calc(formula, en), expected, `English: ${formula}`);
    }
    // What changes changes only in the Turkish locale.
    for (const [formula, expected] of gained) {
      assert.deepEqual(calc(formula, tr), expected, `Turkish, read anew: ${formula}`);
      assert.deepEqual(calc(formula, en), calc(formula, none), `English and no locale: ${formula}`);
      assert.notDeepEqual(calc(formula, none), expected, `no locale, still not read: ${formula}`);
    }
  });
}
// ---------------------------------------------------------------------------
// Without a locale, and in English, text is read exactly as it always was

/** parseNumberText() as it was before a locale was known. */
function oldParseNumberText(text: string, context: CalcContext | undefined): number | undefined {
  const plain = textToNumber(text);
  if (plain !== undefined) return plain;
  if (text.length > 64 || text.length === 0) return undefined;
  const c = text.charCodeAt(0);
  if (c === 61 || c === 39) return undefined;
  const fraction = /^[\s ]*([+-]?)(\d+) +(\d+)\/(\d+)[\s ]*$/.exec(text);
  if (fraction && Number(fraction[4]) !== 0) {
    const value = Number(fraction[2]) + Number(fraction[3]) / Number(fraction[4]);
    if (isFinite(value)) return fraction[1] === '-' ? -value : value;
  }
  let options: { date1904: boolean; today?: Date } = { date1904: false };
  if (context) {
    const year = serialToDate(context.now(), context.date1904).year;
    options = year !== year ? { date1904: context.date1904 === true } : { date1904: context.date1904 === true, today: new Date(Date.UTC(year, 6, 1, 12)) };
  }
  const parsed = parseCellInput(text, options);
  return parsed.kind === 'number' && typeof parsed.value === 'number' && isFinite(parsed.value) ? parsed.value : undefined;
}

function generator(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return state / 4294967296;
  };
}

const PIECES = ['0', '1', '2', '5', '7', '9', '12', '234', '500', '1234', ',', ',', '.', '.', ' ', '-', '+', '%', '$', '(', ')', 'E', 'e', '/', ':', 'x', 'Mar', '₺', ' '];

function randomText(random: () => number): string {
  const count = 1 + Math.floor(random() * 7);
  let text = '';
  for (let i = 0; i < count; i++) text += PIECES[Math.floor(random() * PIECES.length)];
  return text;
}

function contextOf(locale: FormatLocale | undefined): CalcContext {
  return { date1904: false, now: () => 45356.5, locale } as unknown as CalcContext;
}

test('without a locale and in English every text reads as a number exactly as before', () => {
  const random = generator(20261012);
  const none = contextOf(undefined);
  const english = contextOf(LOCALE_EN);
  let numbers = 0;
  for (let i = 0; i < 20000; i++) {
    const text = randomText(random);
    const expected = oldParseNumberText(text, none);
    if (expected !== undefined) numbers++;
    assert.equal(parseNumberText(text, none), expected, `no locale: ${JSON.stringify(text)}`);
    assert.equal(parseNumberText(text, english), expected, `English: ${JSON.stringify(text)}`);
    assert.equal(parseNumberText(text), oldParseNumberText(text, undefined), `no context: ${JSON.stringify(text)}`);
  }
  assert.equal(numbers > 2000, true, `the texts tried hold numbers too (${numbers})`);
});

test('without a locale and in English formulas over random text give what they gave', () => {
  const random = generator(977);
  const cells: Cells = {};
  const texts: string[] = [];
  for (let i = 0; i < 40; i++) {
    const text = randomText(random);
    texts.push(text);
    cells['A' + (i + 1)] = random() < 0.5 ? text : Math.round(random() * 2000) / 4;
  }
  const none = engineOf('none', cells);
  const english = engineOf(LOCALE_EN, cells);
  const quoted = (text: string) => '"' + text.replace(/"/g, '""') + '"';
  for (let i = 0; i < 400; i++) {
    const text = texts[Math.floor(random() * texts.length)];
    const op = ['>', '<', '>=', '<=', '=', '<>', ''][Math.floor(random() * 7)];
    const formulas = [
      `COUNTIF(A1:A40,${quoted(op + text)})`,
      `SUMIF(A1:A40,${quoted(op + text)})`,
      `IFERROR(VALUE(${quoted(text)}),"e")`,
      `IFERROR(${quoted(text)}+1,"e")`,
      `IFERROR(NUMBERVALUE(${quoted(text)}),"e")`,
      `IFERROR(TEXT(${quoted(text)},"#,##0.00"),"e")`,
      `IFERROR(FIXED(${quoted(text)},1),"e")`,
    ];
    for (const formula of formulas) assert.deepEqual(calc(formula, english), calc(formula, none), formula);
  }
});

test('text that is a number or a date the way files write it has the same value in a Turkish locale', () => {
  const random = generator(31337);
  const none = contextOf(undefined);
  const turkish = contextOf(LOCALE_TR);
  let numbers = 0;
  let gained = 0;
  for (let i = 0; i < 60000; i++) {
    const text = randomText(random);
    const expected = oldParseNumberText(text, none);
    const actual = parseNumberText(text, turkish);
    if (expected === undefined) {
      if (actual !== undefined) {
        gained++;
        // What is gained holds the decimal comma or the thousands point of the locale.
        assert.equal(/[.,]/.test(text), true, `read anew without a separator: ${JSON.stringify(text)}`);
      }
      continue;
    }
    numbers++;
    assert.equal(Object.is(actual, expected), true, `Turkish: ${JSON.stringify(text)} is ${String(actual)}, was ${String(expected)}`);
  }
  assert.equal(numbers > 6000, true, `the texts tried hold numbers (${numbers})`);
  assert.equal(gained > 50, true, `and texts only the locale reads (${gained})`);
});

test('every number the engine writes as text is read back as that number in a Turkish locale', () => {
  const random = generator(55);
  const turkish = contextOf(LOCALE_TR);
  const tr = engineOf(LOCALE_TR);
  const fixed = [0, 0.125, 1.234, 12.345, 123.456, 1234.567, 0.001, 1.5, 7.5, 1e-7, 1.234e-7, 1e21, 1.234e21, 1e15, 123456789012345, 0.1 + 0.2, -0.125, -1.234, -12.345, -1e-7, -1e21, 1000, 1.000001];
  const check = (x: number) => {
    const text = numberToText(x);
    const expected = Number(text);
    assert.equal(parseNumberText(text, turkish), expected, `${text} read as text`);
    assert.equal(calc(`VALUE(${text}&"")`, tr), expected, `VALUE(${text}&"")`);
    assert.equal(calc(`"${text}"+0`, tr), expected, `"${text}"+0`);
  };
  for (const x of fixed) check(x);
  for (let i = 0; i < 3000; i++) {
    const digits = Math.floor(random() * 1e6);
    const scale = Math.pow(10, Math.floor(random() * 14) - 7);
    check((random() < 0.3 ? -1 : 1) * digits * scale);
    // Three decimals: what a thousands point would look like.
    check(Math.floor(random() * 1000) + Math.floor(random() * 1000) / 1000);
  }
});

test('odd locales and odd text do no harm', () => {
  const odd = { ...LOCALE_TR, decimal: '5', group: '1' };
  const same = { ...LOCALE_TR, decimal: ',', group: ',' };
  const long = { ...LOCALE_TR, decimal: ',,', group: '..' };
  for (const locale of [odd, long]) {
    const engine = engineOf(locale as FormatLocale);
    assert.equal(calc('VALUE("1.5")', engine), 1.5, 'a locale whose separators cannot be used reads as files do');
    assert.equal(calc('"151"+1', engine), 152, 'digits are never separators');
  }
  const engine = engineOf(same as FormatLocale);
  assert.equal(calc('VALUE("1,5")', engine), 1.5, 'the decimal sign wins when both are the same');
  assert.equal(calc('VALUE("1,234")', engine), 1234, 'and what files write keeps its meaning');
  // Long text comes from a cell: a formula is not that long.
  const many = new Array(5000).join('1.234,');
  const tr = engineOf(LOCALE_TR, { A1: many, A2: '1.234,5', A3: 2 });
  assert.deepEqual(calc('VALUE(A1)', tr), VALUE_ERROR, 'long text');
  assert.deepEqual(calc('A1+1', tr), VALUE_ERROR, 'long text met by an operator');
  assert.equal(calc('COUNTIF(A1:A3,">"&A1)', tr), 1, 'long text in a criterion is text: "1.234,5" comes after it');
  assert.equal(calc('COUNTIF(A1:A3,1234.5)', tr), 1, 'long text among the cells of a criterion');
  assert.equal(calc('VALUE("")', tr) instanceof Object, true, 'empty text is an error');
});
