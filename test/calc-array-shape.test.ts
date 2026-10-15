// The functions of Excel 365 that shape arrays (VSTACK, HSTACK, TAKE, DROP, CHOOSEROWS,
// CHOOSECOLS, EXPAND, TOCOL, TOROW, WRAPROWS, WRAPCOLS) and split text (TEXTSPLIT,
// TEXTBEFORE, TEXTAFTER): calculated in a workbook, spilled, typed with Turkish names, and
// saved to an xlsx file with the prefix files have for them.
//
// Expected values are the examples of Excel's own pages for each function where a comment
// says so; the rest follow the rules those pages give. Excel itself was not at hand.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { FUNCTIONS } from '../src/calc/functions/index';
import type { CalcContext, Value } from '../src/calc/values';
import { needsDynamic } from '../src/calc/dynamic';
import { withFuturePrefixes, withoutFuturePrefixes } from '../src/edit/future-names';
import { formulaFromLocal, formulaToLocal } from '../src/edit/formula-locale';
import { LOCALE_EN, LOCALE_TR } from '../src/format/locale';
import { createBlank } from '../src/formats/blank';
import { writeXlsx } from '../src/formats/xlsx/write';
import { ZipArchive, open } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { createSheet, emptyStyleSheet, getCell, MAX_ROWS, parseCellRef, parseRange, type Cell, type CellFormula, type Row, type Sheet, type Workbook } from '../src/model/sheet/types';

type Plain = number | string | boolean | null;
type Shown = Plain | undefined;

const emptyRow = (): Row => ({ customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] });

/** "=..." is a formula typed into the cell, dynamic where needsDynamic() says so. */
function makeCell(spec: Plain): Cell {
  if (typeof spec === 'string' && spec.charAt(0) === '=') {
    const formula: CellFormula = { text: spec.slice(1), type: 'normal' };
    if (needsDynamic(spec.slice(1))) formula.dynamic = true;
    return { kind: 'blank', value: null, style: 0, formula };
  }
  return { kind: spec === null ? 'blank' : typeof spec === 'number' ? 'number' : typeof spec === 'boolean' ? 'boolean' : 'string', value: spec, style: 0 };
}

function place(sheet: Sheet, address: string, cell: Cell | undefined) {
  const at = parseCellRef(address) as { row: number; col: number };
  const old = sheet.rows.get(at.row);
  const row: Row = old ? { ...old, cells: old.cells.slice() } : emptyRow();
  if (cell) row.cells[at.col] = cell;
  else delete row.cells[at.col];
  sheet.rows.set(at.row, row);
  if (at.row >= sheet.rowCount) sheet.rowCount = at.row + 1;
  if (at.col >= sheet.colCount) sheet.colCount = at.col + 1;
  return { sheet: 0, row: at.row, col: at.col };
}

function book(cells: Record<string, Plain>): Workbook {
  const workbook: Workbook = { sheets: [], styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 };
  const sheet = createSheet({ name: 'S', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  for (const address of Object.keys(cells)) place(sheet, address, makeCell(cells[address]));
  workbook.sheets.push(sheet);
  return workbook;
}

function shown(workbook: Workbook, row: number, col: number): Shown {
  const cell = getCell(workbook.sheets[0], row, col);
  if (!cell) return undefined;
  if (cell.kind === 'error') return `${cell.value}`;
  return cell.value;
}

function calculate(cells: Record<string, Plain>): Workbook {
  const workbook = book(cells);
  const engine = new CalcEngine(workbook, { budget: { milliseconds: 0 } });
  engine.recalculate(engine.uncalculated());
  return workbook;
}

/** What a formula typed into Z1 spills, row by row (the cells its anchor says it fills). */
function spill(formula: string, cells: Record<string, Plain> = {}): Shown[][] {
  const workbook = calculate({ ...cells, Z1: `=${formula}` });
  const anchor = getCell(workbook.sheets[0], 0, 25) as Cell;
  const range = parseRange(anchor.formula && anchor.formula.ref ? anchor.formula.ref : 'Z1')!;
  const out: Shown[][] = [];
  for (let r = range.top; r <= range.bottom; r++) {
    const line: Shown[] = [];
    for (let c = range.left; c <= range.right; c++) line.push(shown(workbook, r, c));
    out.push(line);
  }
  return out;
}

/** The single value of a formula. */
const value = (formula: string, cells: Record<string, Plain> = {}): Shown => {
  const result = spill(formula, cells);
  assert.equal(result.length * result[0].length, 1, `${formula} gives one value: ${JSON.stringify(result)}`);
  return result[0][0];
};

/** 1 to 9 in A1:C3 (the array of the examples of TAKE and DROP). */
const NINE: Record<string, Plain> = { A1: 1, B1: 2, C1: 3, A2: 4, B2: 5, C2: 6, A3: 7, B3: 8, C3: 9 };
const NA = '#N/A';

// ---------------------------------------------------------------------------
// VSTACK and HSTACK

test('VSTACK puts arrays one under the other, HSTACK side by side; what is missing is #N/A', () => {
  // The example of Excel's page: two arrays of three columns.
  const letters = { A1: 'A', B1: 'B', C1: 'C', A2: 'D', B2: 'E', C2: 'F', E1: 'AA', F1: 'BB', G1: 'CC', E2: 'DD', F2: 'EE', G2: 'FF' };
  assert.deepEqual(spill('VSTACK(A1:C2,E1:G2)', letters), [['A', 'B', 'C'], ['D', 'E', 'F'], ['AA', 'BB', 'CC'], ['DD', 'EE', 'FF']]);
  assert.deepEqual(spill('HSTACK(A1:C2,E1:G2)', letters), [['A', 'B', 'C', 'AA', 'BB', 'CC'], ['D', 'E', 'F', 'DD', 'EE', 'FF']]);
  // Of different sizes: padded with #N/A.
  assert.deepEqual(spill('VSTACK(A1:C1,E1:F2)', letters), [['A', 'B', 'C'], ['AA', 'BB', NA], ['DD', 'EE', NA]]);
  assert.deepEqual(spill('HSTACK(A1:A2,{1;2;3})', letters), [['A', 1], ['D', 2], [NA, 3]]);
  // Single values and array constants; an error is stacked as a value.
  assert.deepEqual(spill('VSTACK(1,"x",{2,3},1/0)'), [[1, NA], ['x', NA], [2, 3], ['#DIV/0!', NA]]);
  assert.deepEqual(spill('HSTACK(1,{2;3})'), [[1, 2], [NA, 3]]);
  assert.deepEqual(spill('VSTACK({1,2})'), [[1, 2]]);
  // A range that is no range is the error it is.
  assert.equal(value('ROWS(VSTACK(A1:A2,B1:B3))'), 5);
});

// ---------------------------------------------------------------------------
// TAKE and DROP

test('TAKE: rows and columns from the start, or from the end when negative (the examples of Excel\'s page)', () => {
  assert.deepEqual(spill('TAKE(A1:C3,2)', NINE), [[1, 2, 3], [4, 5, 6]]);
  assert.deepEqual(spill('TAKE(A1:C3,,2)', NINE), [[1, 2], [4, 5], [7, 8]]);
  assert.deepEqual(spill('TAKE(A1:C3,-2)', NINE), [[4, 5, 6], [7, 8, 9]]);
  assert.deepEqual(spill('TAKE(A1:C3,2,2)', NINE), [[1, 2], [4, 5]]);
  assert.deepEqual(spill('TAKE(A1:C3,-1,-1)', NINE), [[9]]);
  // More than there is: all of it. A fraction is cut towards zero.
  assert.deepEqual(spill('TAKE(A1:C3,10,-10)', NINE), [[1, 2, 3], [4, 5, 6], [7, 8, 9]]);
  assert.deepEqual(spill('TAKE(A1:C3,1.9)', NINE), [[1, 2, 3]]);
  assert.deepEqual(spill('TAKE(A1:C3,-1.9)', NINE), [[7, 8, 9]]);
  // None: an empty array is #CALC!.
  assert.equal(value('TAKE(A1:C3,0)', NINE), '#CALC!');
  assert.equal(value('TAKE(A1:C3,2,0)', NINE), '#CALC!');
  assert.equal(value('TAKE(A1:C3,0.5)', NINE), '#CALC!');
  assert.equal(value('TAKE(A1:C3,"x")', NINE), '#VALUE!');
  assert.equal(value('TAKE(A1:C3,1/0)', NINE), '#DIV/0!');
  assert.equal(value('TAKE(A1:C3,{1,2})', NINE), '#VALUE!');
  // The rows of an expression, a single value.
  assert.deepEqual(spill('TAKE(SEQUENCE(5)*10,-2)'), [[40], [50]]);
  assert.deepEqual(spill('TAKE(7,1)'), [[7]]);
});

test('DROP: leaves out rows and columns from the start, or from the end when negative (the examples of Excel\'s page)', () => {
  assert.deepEqual(spill('DROP(A1:C3,2)', NINE), [[7, 8, 9]]);
  assert.deepEqual(spill('DROP(A1:C3,,2)', NINE), [[3], [6], [9]]);
  assert.deepEqual(spill('DROP(A1:C3,-2)', NINE), [[1, 2, 3]]);
  assert.deepEqual(spill('DROP(A1:C3,2,2)', NINE), [[9]]);
  assert.deepEqual(spill('DROP(A1:C3,-1,-2)', NINE), [[1], [4]]);
  // Everything dropped: #CALC!.
  assert.equal(value('DROP(A1:C3,3)', NINE), '#CALC!');
  assert.equal(value('DROP(A1:C3,-5)', NINE), '#CALC!');
  assert.equal(value('DROP(A1:C3,1,3)', NINE), '#CALC!');
  // UNSURE: none dropped for 0 (Excel's page says #CALC! for 0, with the words of TAKE).
  assert.deepEqual(spill('DROP(A1:C3,0,1)', NINE), [[2, 3], [5, 6], [8, 9]]);
});

// ---------------------------------------------------------------------------
// CHOOSEROWS and CHOOSECOLS

test('CHOOSEROWS and CHOOSECOLS pick rows and columns by number, arrays of numbers, from the end when negative', () => {
  const table: Record<string, Plain> = {};
  for (let r = 1; r <= 6; r++) {
    table[`A${r}`] = r;
    table[`B${r}`] = r * 10;
  }
  // The examples of Excel's pages.
  assert.deepEqual(spill('CHOOSEROWS(A1:B6,1,3,5,1)', table), [[1, 10], [3, 30], [5, 50], [1, 10]]);
  assert.deepEqual(spill('CHOOSEROWS(A1:B6,{1,3},5,1)', table), [[1, 10], [3, 30], [5, 50], [1, 10]]);
  assert.deepEqual(spill('CHOOSEROWS(A1:B6,-1,-2)', table), [[6, 60], [5, 50]]);
  assert.deepEqual(spill('CHOOSECOLS(A1:C3,1,3,1)', NINE), [[1, 3, 1], [4, 6, 4], [7, 9, 7]]);
  assert.deepEqual(spill('CHOOSECOLS(A1:C3,-1)', NINE), [[3], [6], [9]]);
  assert.deepEqual(spill('CHOOSECOLS(A1:C3,SEQUENCE(1,2,3,-1))', NINE), [[3, 2], [6, 5], [9, 8]]);
  // 0, or beyond the array either way: #VALUE!.
  assert.equal(value('CHOOSEROWS(A1:B6,0)', table), '#VALUE!');
  assert.equal(value('CHOOSEROWS(A1:B6,7)', table), '#VALUE!');
  assert.equal(value('CHOOSEROWS(A1:B6,-7)', table), '#VALUE!');
  assert.equal(value('CHOOSECOLS(A1:C3,4)', NINE), '#VALUE!');
  assert.equal(value('CHOOSECOLS(A1:C3,"a")', NINE), '#VALUE!');
  assert.equal(value('CHOOSECOLS(A1:C3,1,NA())', NINE), NA);
  assert.deepEqual(spill('CHOOSEROWS(A1:C3,2.7)', NINE), [[4, 5, 6]]);
});

// ---------------------------------------------------------------------------
// EXPAND

test('EXPAND makes an array larger, padded with #N/A or what is given', () => {
  const four = { A1: 1, B1: 2, A2: 3, B2: 4 };
  assert.deepEqual(spill('EXPAND(A1:B2,3,3)', four), [[1, 2, NA], [3, 4, NA], [NA, NA, NA]]);
  assert.deepEqual(spill('EXPAND(A1:B2,3,3,"-")', four), [[1, 2, '-'], [3, 4, '-'], ['-', '-', '-']]);
  assert.deepEqual(spill('EXPAND(A1:B2,,3,0)', four), [[1, 2, 0], [3, 4, 0]]);
  assert.deepEqual(spill('EXPAND(A1:B2,3)', four), [[1, 2], [3, 4], [NA, NA]]);
  assert.deepEqual(spill('EXPAND(A1:B2,2,2)', four), [[1, 2], [3, 4]]);
  // Smaller than the array: #VALUE!.
  assert.equal(value('EXPAND(A1:B2,1)', four), '#VALUE!');
  assert.equal(value('EXPAND(A1:B2,2,1)', four), '#VALUE!');
  assert.equal(value('EXPAND(A1:B2,-3)', four), '#VALUE!');
  // Too large to make at all: #NUM!, at once, without making it.
  const started = Date.now();
  assert.equal(value('EXPAND(A1:B2,1000000,1000000)', four), '#NUM!');
  assert.equal(value('EXPAND(A1:B2,1E+300,2)', four), '#NUM!');
  assert.ok(Date.now() - started < 2000, `${Date.now() - started} ms`);
  // Larger than a sheet holds but small enough to make: it does not spill.
  assert.equal(value('EXPAND(A1:B2,2000000)', four), '#SPILL!');
});

// ---------------------------------------------------------------------------
// TOCOL and TOROW

test('TOCOL and TOROW lay an array out in one column or row, by rows or by columns, leaving out blanks and errors when asked', () => {
  // The example of Excel's page, with two blank cells and an error put in.
  const names: Record<string, Plain> = { A1: 'Ben', B1: 'Peter', C1: 'Mary', A2: 'John', B2: '=1/0', C2: 'Jenny', A3: 'Agnes', B3: null };
  assert.deepEqual(spill('TOCOL(A1:C3,3)', names), [['Ben'], ['Peter'], ['Mary'], ['John'], ['Jenny'], ['Agnes']]);
  assert.deepEqual(spill('TOCOL(A1:C3,3,TRUE)', names), [['Ben'], ['John'], ['Agnes'], ['Peter'], ['Mary'], ['Jenny']]);
  assert.deepEqual(spill('TOROW(A1:C3,1)', names), [['Ben', 'Peter', 'Mary', 'John', '#DIV/0!', 'Jenny', 'Agnes']]);
  assert.deepEqual(spill('TOROW(A1:C3,2)', names)[0].length, 8);
  assert.equal(spill('TOCOL(A1:C3)', names).length, 9);
  assert.deepEqual(spill('TOCOL(A1:C3,0,1)', NINE), [[1], [4], [7], [2], [5], [8], [3], [6], [9]]);
  assert.deepEqual(spill('TOROW({1,2;3,4})'), [[1, 2, 3, 4]]);
  // Everything left out: #CALC!; an ignore that is not 0 to 3: #VALUE!.
  assert.equal(value('TOCOL(B3:C3,1)', names), '#CALC!');
  assert.equal(value('TOCOL(A1:C3,4)', names), '#VALUE!');
  assert.equal(value('TOCOL(A1:C3,-1)', names), '#VALUE!');
  assert.equal(value('TOCOL(A1:C3,0,"x")', names), '#VALUE!');
});

// ---------------------------------------------------------------------------
// WRAPROWS and WRAPCOLS

test('WRAPROWS and WRAPCOLS fold a row or column into rows or columns of so many values (the examples of Excel\'s pages)', () => {
  const seven = { A1: 'A', B1: 'B', C1: 'C', D1: 'D', E1: 'E', F1: 'F', G1: 'G' };
  assert.deepEqual(spill('WRAPROWS(A1:G1,3)', seven), [['A', 'B', 'C'], ['D', 'E', 'F'], ['G', NA, NA]]);
  assert.deepEqual(spill('WRAPROWS(A1:G1,3,"x")', seven), [['A', 'B', 'C'], ['D', 'E', 'F'], ['G', 'x', 'x']]);
  assert.deepEqual(spill('WRAPCOLS(A1:G1,3)', seven), [['A', 'D', 'G'], ['B', 'E', NA], ['C', 'F', NA]]);
  assert.deepEqual(spill('WRAPCOLS(A1:G1,3,"x")', seven), [['A', 'D', 'G'], ['B', 'E', 'x'], ['C', 'F', 'x']]);
  assert.deepEqual(spill('WRAPROWS(SEQUENCE(4),2)'), [[1, 2], [3, 4]]);
  assert.deepEqual(spill('WRAPCOLS(SEQUENCE(1,4),4)'), [[1], [2], [3], [4]]);
  // UNSURE: a count larger than the vector still makes a line of that many.
  assert.deepEqual(spill('WRAPROWS({1,2},3)'), [[1, 2, NA]]);
  // Not a vector: #VALUE!; a count below 1: #NUM!.
  assert.equal(value('WRAPROWS(A1:C3,2)', NINE), '#VALUE!');
  assert.equal(value('WRAPCOLS(A1:G1,0)', seven), '#NUM!');
  assert.equal(value('WRAPROWS(A1:G1,1E+300)', seven), '#NUM!');
});

// ---------------------------------------------------------------------------
// TEXTSPLIT, TEXTBEFORE, TEXTAFTER

test('TEXTSPLIT splits at column and row delimiters, several of them, and pads the rows (the examples of Excel\'s page)', () => {
  assert.deepEqual(spill('TEXTSPLIT("Dakota Lennox"," ")'), [['Dakota', 'Lennox']]);
  assert.deepEqual(spill('TEXTSPLIT("Do. Or do not. There is no try. -Anonymous",".")'), [['Do', ' Or do not', ' There is no try', ' -Anonymous']]);
  assert.deepEqual(spill('TEXTSPLIT("Do. Or do not. There is no try. -Anonymous",,".")'), [['Do'], [' Or do not'], [' There is no try'], [' -Anonymous']]);
  assert.deepEqual(spill('TEXTSPLIT("Do. Or do not. There is no try. -Anonymous",{".","-"},,TRUE)'), [['Do', ' Or do not', ' There is no try', ' ', 'Anonymous']]);
  assert.deepEqual(spill('TEXTSPLIT("1,2,3;4,5,6",",",";")'), [['1', '2', '3'], ['4', '5', '6']]);
  // Rows of fewer pieces are padded with #N/A, or with what is given.
  assert.deepEqual(spill('TEXTSPLIT("1,2,3;4;5,6",",",";")'), [['1', '2', '3'], ['4', NA, NA], ['5', '6', NA]]);
  assert.deepEqual(spill('TEXTSPLIT("1,2,3;4;5,6",",",";",,,"-")'), [['1', '2', '3'], ['4', '-', '-'], ['5', '6', '-']]);
  // Empty pieces stay unless ignore_empty; empty rows too.
  assert.deepEqual(spill('TEXTSPLIT("a,,b",",")'), [['a', '', 'b']]);
  assert.deepEqual(spill('TEXTSPLIT("a,,b",",",,TRUE)'), [['a', 'b']]);
  assert.deepEqual(spill('TEXTSPLIT(",a;;b,",",",";",TRUE)'), [['a'], ['b']]);
  assert.deepEqual(spill('TEXTSPLIT("a;",",",";")'), [['a'], ['']]);
  // match_mode 1: without regard to case.
  assert.deepEqual(spill('TEXTSPLIT("aXbxc","x")'), [['aXb', 'c']]);
  assert.deepEqual(spill('TEXTSPLIT("aXbxc","x",,,1)'), [['a', 'b', 'c']]);
  assert.equal(value('TEXTSPLIT("a","x",,,2)'), '#VALUE!');
  // Delimiters from a range; the longer of two found at one place.
  assert.deepEqual(spill('TEXTSPLIT("a--b-c",A1:A2)', { A1: '-', A2: '--' }), [['a', 'b', 'c']]);
  // No delimiter found: the text; nothing left: #CALC!; no delimiter given: #VALUE!.
  assert.equal(value('TEXTSPLIT("abc",",")'), 'abc');
  assert.equal(value('TEXTSPLIT("",",")'), '');
  assert.equal(value('TEXTSPLIT(",,",",",,TRUE)'), '#CALC!');
  assert.equal(value('TEXTSPLIT("abc","")'), '#VALUE!');
  assert.equal(value('TEXTSPLIT(1/0,",")'), '#DIV/0!');
  assert.equal(value('TEXTSPLIT("a,b",NA())'), NA);
  // Numbers are split as their text, and the pieces are text.
  assert.deepEqual(spill('TEXTSPLIT(1.5,".")'), [['1', '5']]);
  // UNSURE: several texts give the first piece of each, as an array of arrays does in Excel.
  assert.deepEqual(spill('TEXTSPLIT(A1:A2,",")', { A1: 'a,b', A2: 'c,d' }), [['a'], ['c']]);
});

test('TEXTSPLIT of a long text takes time in proportion to its length, however many delimiters there are', () => {
  // One delimiter is everywhere, the others nowhere, or only at the very end: no piece searches the rest of the text again.
  const pieces = 30000;
  const long = 'ab,'.repeat(pieces) + ';end';
  const started = Date.now();
  const result = spill('ROWS(TEXTSPLIT(A1,{",","|","#"},";"))&"/"&COLUMNS(TEXTSPLIT(A1,{",","|","#"},";"))', { A1: long });
  const took = Date.now() - started;
  assert.deepEqual(result, [[`2/${pieces + 1}`]]);
  assert.ok(took < 5000, `${took} ms`);
  // Too many pieces in rows and columns to make: #NUM!, not an array of a billion cells.
  const square = ','.repeat(30000) + ';'.repeat(30000);
  assert.equal(value('TEXTSPLIT(A1,",",";")', { A1: square }), '#NUM!');
});

test('TEXTBEFORE and TEXTAFTER take several delimiters, count from the end, match the end, and give if_not_found', () => {
  const hood = { A1: "Little red riding hood's red hood" };
  // The examples of Excel's pages.
  assert.equal(value('TEXTBEFORE(A1,"hood")', hood), 'Little red riding ');
  assert.equal(value('TEXTBEFORE(A1,"red",2)', hood), "Little red riding hood's ");
  assert.equal(value('TEXTBEFORE(A1,"red",-2)', hood), 'Little ');
  assert.equal(value('TEXTBEFORE(A1,"Red")', hood), NA);
  assert.equal(value('TEXTBEFORE(A1,"Red",,1)', hood), 'Little ');
  assert.equal(value('TEXTAFTER(A1,"hood")', hood), "'s red hood");
  assert.equal(value('TEXTAFTER(A1,"red",2)', hood), ' hood');
  assert.equal(value('TEXTAFTER(A1,"red",-2)', hood), " riding hood's red hood");
  assert.equal(value('TEXTAFTER(A1,"basket",,,,"none")', hood), 'none');
  assert.equal(value('TEXTBEFORE("Socrates"," ",,,1)'), 'Socrates');
  assert.equal(value('TEXTAFTER("Socrates"," ",-1,,1)'), 'Socrates');
  // Several delimiters: whichever comes first (from the end when counting back).
  assert.equal(value('TEXTBEFORE("a,b;c",{";",","})'), 'a');
  assert.equal(value('TEXTBEFORE("a,b;c",{";",","},2)'), 'a,b');
  assert.equal(value('TEXTAFTER("a,b;c",{";",","},-1)'), 'c');
  assert.equal(value('TEXTAFTER("a,b;c",{";",","},-2)'), 'b;c');
  assert.equal(value('TEXTBEFORE("a,b;c",B1:B2)', { B1: ';', B2: ',' }), 'a');
  assert.equal(value('TEXTBEFORE("aXbxc",{"y","X"},2,1)'), 'aXb');
  assert.equal(value('TEXTBEFORE("a,b",{"x","y"})'), NA);
  assert.equal(value('TEXTBEFORE("a,b",B1:B2)', { B1: 'x', B2: '=1/0' }), '#DIV/0!');
  // The texts of a range each give theirs.
  assert.deepEqual(spill('TEXTBEFORE(A1:A3,"-")', { A1: 'a-b', A2: 'cc-d', A3: 'e' }), [['a'], ['cc'], [NA]]);
  assert.deepEqual(spill('TEXTAFTER("a-b-c","-",{1,2})'), [['b-c', 'c']]);
  // instance_num 0 or beyond the length: #VALUE!.
  assert.equal(value('TEXTBEFORE("a-b","-",0)'), '#VALUE!');
  assert.equal(value('TEXTBEFORE("a-b","-",4)'), '#VALUE!');
});

/** 1,024 delimiters of one character each, and texts in which the last of them stands every other character. */
function manyDelimiters(texts: number, length: number): Record<string, Plain> {
  const cells: Record<string, Plain> = {};
  const ch = (i: number) => String.fromCharCode(0x4e00 + i);
  for (let i = 0; i < 1024; i++) {
    cells[`B${i + 1}`] = ch(i);
    cells[`C${i + 1}`] = ch(1024 + i);
  }
  let text = '';
  while (text.length < length) text += 'a' + ch(1023);
  for (let r = 1; r <= texts; r++) cells[`A${r}`] = text.slice(0, length);
  return cells;
}

/** The value of a formula in Z1 calculated with a budget of so many milliseconds a call, and how long it took. */
function timed(formula: string, cells: Record<string, Plain>, milliseconds: number): { value: Shown; took: number } {
  const workbook = book({ ...cells, Z1: `=${formula}` });
  const engine = new CalcEngine(workbook, { budget: { milliseconds } });
  const started = Date.now();
  engine.recalculate(engine.uncalculated());
  for (let k = 0; k < 1000 && engine.incomplete; k++) engine.recalculate([]);
  return { value: shown(workbook, 0, 25), took: Date.now() - started };
}

test('many delimiters in many long texts: the work grows with the text, not with the delimiters times the pieces (review B1)', () => {
  // 300 texts of 32,766 characters, 16,383 delimiters in each, 1,024 delimiters to look for.
  const cells = manyDelimiters(300, 32766);
  const before = timed('SUM(LEN(TEXTBEFORE(A1:A300,B1:B1024,-16000)))', cells, 250);
  // TEXTBEFORE from the end: the 16,000th delimiter from the end is the 384th from the start, at 767.
  assert.equal(before.value, 300 * 767);
  assert.ok(before.took < 5000, `TEXTBEFORE ${before.took} ms`);
  const split = timed('ROWS(TEXTSPLIT(A1,B1:B1024,C1:C1024))&"/"&COLUMNS(TEXTSPLIT(A1,B1:B1024,C1:C1024))', cells, 250);
  assert.equal(split.value, '1/16384');
  assert.ok(split.took < 5000, `TEXTSPLIT ${split.took} ms`);
});

test('a formula of many delimiters that takes long keeps to the budget of a call (review B1)', () => {
  // The worst there is for the search: 1,024 delimiters that all begin with "a", none found, in texts of nothing but "a".
  // Each place of each text is looked at with every one of them. However long that takes in all, no one call takes much
  // longer than its budget (ten times that, the hard limit, for a formula that cannot be cut), and the formula is then
  // given up as one that takes too long (CalcEngine.overBudget) instead of hanging the page.
  const cells: Record<string, Plain> = {};
  for (let i = 0; i < 1024; i++) cells[`B${i + 1}`] = 'a' + String.fromCharCode(0x4e00 + i);
  const text = 'a'.repeat(32766);
  for (let r = 1; r <= 2000; r++) cells[`A${r}`] = text;
  const workbook = book({ ...cells, Z1: '=SUM(LEN(TEXTBEFORE(A1:A2000,B1:B1024,1,,,"")))' });
  const engine = new CalcEngine(workbook, { budget: { milliseconds: 100 } });
  for (let call = 0; call < 5; call++) {
    const started = Date.now();
    engine.recalculate(call === 0 ? engine.uncalculated() : []);
    const took = Date.now() - started;
    assert.ok(took < 3000, `call ${call} took ${took} ms`);
    if (!engine.incomplete) break;
  }
  assert.equal(engine.incomplete, false);
  assert.deepEqual(engine.overBudget, [{ sheet: 0, row: 0, col: 25 }]);
});

test('TEXTSPLIT of several texts splits each only as far as its first piece (review B2)', () => {
  const cells: Record<string, Plain> = {};
  const text = ',' + 'a'.repeat(32766);
  for (let r = 1; r <= 2000; r++) cells[`A${r}`] = text;
  const workbook = calculate({ ...cells, C1: '=TEXTSPLIT(A1:A2000,",")', D1: '=COUNTIF(C1:C2000,"")' });
  // Each gives "", the piece before its comma; no text of 32,766 characters is made 2,000 times over.
  assert.equal(shown(workbook, 1999, 2), '');
  assert.equal(shown(workbook, 0, 3), 2000);
  // With ignore_empty the first piece is the first that is not empty.
  assert.deepEqual(spill('TEXTSPLIT(A1:A2,",",";",TRUE)', { A1: ',,a,b', A2: ';;' }), [['a'], ['#CALC!']]);
  assert.deepEqual(spill('TEXTSPLIT(A1:A2,",",";")', { A1: 'x;y,z', A2: 'p,q' }), [['x'], ['p']]);
});

test('a delimiter left empty is an empty text, as an empty cell is (review B4; UNSURE: Excel not checked)', () => {
  assert.equal(value('TEXTBEFORE("abc",)'), '');
  assert.equal(value('TEXTAFTER("abc",)'), 'abc');
  assert.equal(value('TEXTBEFORE("abc",B1)'), '');
  assert.equal(value('TEXTAFTER("abc",B1)'), 'abc');
  // UNSURE (review B5, B7): an empty delimiter is found at once here and passed over by TEXTSPLIT; counting from the end
  // takes the delimiter that starts last, from the start the one that starts first (and is longest).
  assert.equal(value('TEXTBEFORE("abc","")'), '');
  assert.deepEqual(spill('TEXTSPLIT("abc",{"","b"})'), [['a', 'c']]);
  assert.equal(value('TEXTAFTER("a--b",{"-","--"})'), 'b');
  assert.equal(value('TEXTAFTER("a--b",{"-","--"},-1)'), 'b');
  assert.equal(value('TEXTBEFORE("a--b",{"-","--"},-1)'), 'a-');
});

test('VSTACK and HSTACK of references too large in all are #NUM! before any cell is read (review B6)', () => {
  const rows = 250000;
  let visited = 0;
  const context: CalcContext = {
    cellValue: () => 1,
    forEachCell(_sheet, range, visit) {
      for (let r = range.top; r <= Math.min(range.bottom, rows - 1); r++) {
        visited++;
        visit(r, r, range.left);
      }
    },
    usedExtent: () => ({ rows, cols: 1 }),
    sheetIndex: () => 0,
    sheetName: () => 'S',
    sheetCount: () => 1,
    sheet: 0,
    row: 0,
    col: 30,
    date1904: false,
    now: () => 0,
    random: () => 0.5,
  };
  const column: Value = { ref: true, sheet: 0, range: { top: 0, left: 0, bottom: MAX_ROWS - 1, right: 0 } };
  const args = new Array(254).fill(column);
  for (const name of ['VSTACK', 'HSTACK', 'CHOOSEROWS', 'CHOOSECOLS']) {
    visited = 0;
    const spec = FUNCTIONS.get(name)!;
    const result = spec.call(name.startsWith('CHOOSE') ? [column, ...new Array(253).fill(column)] : args, context);
    if (name === 'VSTACK' || name === 'HSTACK') assert.deepEqual(result, { error: '#NUM!' }, name);
    assert.ok(visited <= rows, `${name} read ${visited} cells`);
  }
});

test('TEXTBEFORE of a long text with delimiters that are not there takes no longer than one search for each', () => {
  const long = 'x'.repeat(200000) + ',y';
  const started = Date.now();
  assert.equal(value('LEN(TEXTBEFORE(A1,{";","|",","},1))', { A1: long }), 200000);
  assert.equal(value('TEXTAFTER(A1,{";","|",","},-1)', { A1: long }), 'y');
  assert.ok(Date.now() - started < 5000, `${Date.now() - started} ms`);
});

// ---------------------------------------------------------------------------
// Spilling

test('the results spill, and something in the way makes them #SPILL!', () => {
  const cells = { A1: 1, A2: 2, B1: 3, B2: 4, B3: 5 };
  const workbook = calculate({ ...cells, D1: '=VSTACK(A1:A2,B1:B3)', F1: '=TOROW(A1:B3,1)' });
  assert.deepEqual([0, 1, 2, 3, 4, 5].map((row) => shown(workbook, row, 3)), [1, 2, 3, 4, 5, undefined]);
  assert.deepEqual([5, 6, 7, 8, 9, 10].map((col) => shown(workbook, 0, col)), [1, 3, 2, 4, 5, undefined]);
  assert.equal((getCell(workbook.sheets[0], 0, 3) as Cell).formula?.ref, 'D1:D5');

  const blocked = calculate({ ...cells, D1: '=VSTACK(A1:A2,B1:B3)', D4: 'x' });
  assert.equal(shown(blocked, 0, 3), '#SPILL!');
  assert.equal(shown(blocked, 3, 3), 'x');

  // What reads the spill with "#".
  const read = calculate({ ...cells, D1: '=HSTACK(A1:A2,B1:B2)', G1: '=SUM(D1#)', H1: '=ROWS(TAKE(D1#,-1))' });
  assert.equal(shown(read, 0, 6), 10);
  assert.equal(shown(read, 0, 7), 1);
});

test('the functions are dynamic array formulas wherever they stand', () => {
  for (const text of ['VSTACK(A1,B1)', 'SUM(TAKE(A1:A9,2))+0', 'TEXTSPLIT(A1,",")', 'TOCOL(A1:B2)', 'EXPAND(A1,2)', 'WRAPROWS(A1:A4,2)']) {
    assert.equal(needsDynamic(text), true, text);
  }
  // One value from one text and one delimiter.
  assert.equal(needsDynamic('TEXTBEFORE(A1,"-")'), false);
});

// ---------------------------------------------------------------------------
// Names in files and in Turkish

test('files store them with "_xlfn.", shown and typed without it', () => {
  const cases: [shown: string, stored: string][] = [
    ['VSTACK(A1:A2,HSTACK(B1,C1))', '_xlfn.VSTACK(A1:A2,_xlfn.HSTACK(B1,C1))'],
    ['TAKE(DROP(A1:C9,1),-2,CHOOSECOLS(1,1))', '_xlfn.TAKE(_xlfn.DROP(A1:C9,1),-2,_xlfn.CHOOSECOLS(1,1))'],
    ['CHOOSEROWS(EXPAND(A1,2,2,0),1)', '_xlfn.CHOOSEROWS(_xlfn.EXPAND(A1,2,2,0),1)'],
    ['TOCOL(WRAPROWS(TOROW(A1:B2),1))+WRAPCOLS(A1:A4,2)', '_xlfn.TOCOL(_xlfn.WRAPROWS(_xlfn.TOROW(A1:B2),1))+_xlfn.WRAPCOLS(A1:A4,2)'],
    ['TEXTSPLIT(TEXTBEFORE(A1,"-"),TEXTAFTER(A1,","))', '_xlfn.TEXTSPLIT(_xlfn.TEXTBEFORE(A1,"-"),_xlfn.TEXTAFTER(A1,","))'],
    ['LET(x,A1:A3,VSTACK(x,x#))', '_xlfn.LET(_xlpm.x,A1:A3,_xlfn.VSTACK(_xlpm.x,_xlfn.ANCHORARRAY(_xlpm.x)))'],
  ];
  for (const [typed, stored] of cases) {
    assert.equal(withFuturePrefixes(typed), stored, typed);
    assert.equal(withoutFuturePrefixes(stored), typed, stored);
    assert.equal(formulaFromLocal(typed, LOCALE_EN), stored);
    assert.equal(formulaToLocal(stored, LOCALE_EN), typed);
  }
  // A cell or a sheet called like one is no call.
  assert.equal(withFuturePrefixes('TAKE!A1+DROP'), 'TAKE!A1+DROP');
});

test('typed in Turkish (UNSURE: the names are from memory, not from a Turkish Excel)', () => {
  assert.equal(formulaFromLocal('DÜŞEYYIĞ(A1:A2;B1:B2)', LOCALE_TR), '_xlfn.VSTACK(A1:A2,B1:B2)');
  assert.equal(formulaFromLocal('dikeyyığ(A1;2)', LOCALE_TR), '_xlfn.VSTACK(A1,2)');
  assert.equal(formulaFromLocal('YATAYYIĞ(A1;B1)', LOCALE_TR), '_xlfn.HSTACK(A1,B1)');
  assert.equal(formulaFromLocal('AL(A1:C3;2)+BIRAK(A1:C3;-1)', LOCALE_TR), '_xlfn.TAKE(A1:C3,2)+_xlfn.DROP(A1:C3,-1)');
  assert.equal(formulaFromLocal('SATIRSEÇ(A1:C3;1)&SÜTUNSEÇ(A1:C3;2)', LOCALE_TR), '_xlfn.CHOOSEROWS(A1:C3,1)&_xlfn.CHOOSECOLS(A1:C3,2)');
  assert.equal(formulaFromLocal('GENİŞLET(A1;2;2;0,5)', LOCALE_TR), '_xlfn.EXPAND(A1,2,2,0.5)');
  assert.equal(formulaFromLocal('SÜTUNA(A1:B2)&SATIRA(A1:B2)', LOCALE_TR), '_xlfn.TOCOL(A1:B2)&_xlfn.TOROW(A1:B2)');
  assert.equal(formulaFromLocal('SATIRSAR(A1:A4;2)&SÜTUNSAR(A1:A4;2)', LOCALE_TR), '_xlfn.WRAPROWS(A1:A4,2)&_xlfn.WRAPCOLS(A1:A4,2)');
  assert.equal(formulaFromLocal('METİNBÖL(A1;",";";")', LOCALE_TR), '_xlfn.TEXTSPLIT(A1,",",";")');
  assert.equal(formulaFromLocal('ÖNCEKİMETİN(A1;"-")&SONRAKİMETİN(A1;"-")', LOCALE_TR), '_xlfn.TEXTBEFORE(A1,"-")&_xlfn.TEXTAFTER(A1,"-")');
  // Shown with the English names.
  assert.equal(formulaToLocal('_xlfn.TAKE(A1:C3,2,0.5)', LOCALE_TR), 'TAKE(A1:C3;2;0,5)');
  // The column AL is still a column.
  assert.equal(formulaFromLocal('AL1+AL:AL', LOCALE_TR), 'AL1+AL:AL');
});

test('typed into an xlsx file and saved: "_xlfn." in the file, the cells of the spill, and read again the same', async () => {
  const file = await open(await createBlank('xlsx'));
  const workbook = await file.workbook();
  const engine = new CalcEngine(workbook);
  const typed = (row: number, col: number, input: string) => {
    for (const op of setCellInput(workbook, 0, row, col, input)) {
      if (op.type === 'setCells') {
        for (const change of op.cells) {
          const sheet = workbook.sheets[0];
          let line = sheet.rows.get(change.row);
          line = line ? { ...line, cells: line.cells.slice() } : emptyRow();
          if (change.cell) line.cells[change.col] = change.cell;
          sheet.rows.set(change.row, line);
          sheet.changedRows = (sheet.changedRows || new Set()).add(change.row);
          sheet.modified = true;
          if (change.row >= sheet.rowCount) sheet.rowCount = change.row + 1;
          if (change.col >= sheet.colCount) sheet.colCount = change.col + 1;
        }
      }
    }
    engine.recalculate([{ sheet: 0, row, col }]);
  };
  typed(0, 0, '1');
  typed(1, 0, '2');
  typed(0, 1, '3');
  typed(0, 3, '=VSTACK(A1:A2,B1)');
  typed(0, 4, '=TEXTSPLIT("a,b;c",",",";",,,"-")');
  typed(0, 7, '=TEXTBEFORE("x-y","-")');
  await writeXlsx(file.package!, workbook);
  const saved = await file.package!.save();
  const sheet = new TextDecoder().decode(await ZipArchive.read(saved).read('xl/worksheets/sheet1.xml'));
  assert.match(sheet, /<c r="D1" cm="1"><f t="array" ref="D1:D3">_xlfn\.VSTACK\(A1:A2,B1\)<\/f><v>1<\/v><\/c>/);
  assert.match(sheet, /<c r="E1" cm="1" t="str"><f t="array" ref="E1:F2">_xlfn\.TEXTSPLIT\("a,b;c",",",";",,,"-"\)<\/f><v>a<\/v><\/c>/);
  assert.match(sheet, /<c r="H1" t="str"><f>_xlfn\.TEXTBEFORE\("x-y","-"\)<\/f><v>x<\/v><\/c>/);

  const again = await (await open(saved)).workbook();
  const anchor = getCell(again.sheets[0], 0, 3)!;
  assert.equal(anchor.formula?.dynamic, true);
  assert.equal(anchor.formula?.text, '_xlfn.VSTACK(A1:A2,B1)');
  assert.equal(getCell(again.sheets[0], 2, 3)?.value, 3);
  assert.equal(getCell(again.sheets[0], 1, 5)?.value, '-');
  // Calculated again from the file: the same.
  const fresh = new CalcEngine(again, { budget: { milliseconds: 0 } });
  fresh.recalculate([{ sheet: 0, row: 0, col: 0 }]);
  assert.deepEqual([0, 1, 2].map((row) => getCell(again.sheets[0], row, 3)?.value), [1, 2, 3]);
});

// ---------------------------------------------------------------------------
// Size

test('10,000 rows stacked and laid out in a column take little time', () => {
  const cells: Record<string, Plain> = {};
  for (let r = 1; r <= 10000; r++) {
    cells[`A${r}`] = r;
    cells[`B${r}`] = `t${r}`;
  }
  let started = Date.now();
  const stacked = calculate({ ...cells, D1: '=VSTACK(A1:B10000,B1:A10000)' });
  const stack = Date.now() - started;
  assert.equal(shown(stacked, 19999, 3), 10000);
  started = Date.now();
  const column = calculate({ ...cells, D1: '=TOCOL(A1:B10000,1,TRUE)' });
  const tocol = Date.now() - started;
  assert.equal(shown(column, 19999, 3), 't10000');
  // Guards against hanging, not measurements: some ten times what an idle machine takes.
  assert.ok(stack < 10000 && tocol < 10000, `VSTACK ${stack} ms, TOCOL ${tocol} ms`);
});
