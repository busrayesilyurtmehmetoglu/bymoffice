// Tests for the defects an independent review of the formula engine confirmed: implicit
// intersection, number <-> text conversion, nested subtotals, 30/360, formulas the engine
// cannot calculate, circular references, omitted arguments.
//
// There is no Excel on the machine these tests were written on. Expected values come from
// Excel's documented behaviour; what is not certain is marked UNSURE.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine, type CellAddress } from '../src/calc/engine';
import { FUNCTIONS } from '../src/calc/functions/index';
import type { FunctionSpec } from '../src/calc/values';
import { cellName, createSheet, emptyStyleSheet, getCell, parseCellRef, type Cell, type CellFormula, type DefinedName, type Row, type Sheet, type Workbook } from '../src/model/sheet/types';

// ---------------------------------------------------------------------------
// Building workbooks by hand

type Plain = number | string | boolean | null;
interface FormulaSpec {
  f: string;
  /** Value the file would have stored for the formula. */
  v?: Plain;
  type?: CellFormula['type'];
  ref?: string;
  error?: boolean;
  /** Cell metadata index: set by Excel on dynamic array formulas. */
  cm?: number;
}
type Spec = Plain | FormulaSpec;
type Cells = Record<string, Spec>;

function kindOf(value: Plain, error = false): Cell['kind'] {
  if (error) return 'error';
  return value === null ? 'blank' : typeof value === 'number' ? 'number' : typeof value === 'boolean' ? 'boolean' : 'string';
}

/** A string starting with "=" is a formula that was never calculated. */
function makeCell(spec: Spec): Cell {
  if (typeof spec === 'string' && spec.charAt(0) === '=') spec = { f: spec.slice(1) };
  if (spec !== null && typeof spec === 'object') {
    const value = spec.v === undefined ? null : spec.v;
    const formula: CellFormula = { text: spec.f, type: spec.type || 'normal' };
    if (spec.ref !== undefined) formula.ref = spec.ref;
    const cell: Cell = { kind: kindOf(value, spec.error), value, style: 0, formula };
    if (spec.cm !== undefined) cell.cm = spec.cm;
    return cell;
  }
  return { kind: kindOf(spec), value: spec, style: 0 };
}

const emptyRow = (): Row => ({ customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] });

function place(sheet: Sheet, row: number, col: number, cell: Cell): void {
  let target = sheet.rows.get(row);
  if (!target) sheet.rows.set(row, (target = emptyRow()));
  target.cells[col] = cell;
  if (row >= sheet.rowCount) sheet.rowCount = row + 1;
  if (col >= sheet.colCount) sheet.colCount = col + 1;
}

function book(sheets: Record<string, Cells>, definedNames: DefinedName[] = []): Workbook {
  const workbook: Workbook = { sheets: [], styles: emptyStyleSheet(), sharedStrings: [], definedNames, date1904: false, activeSheet: 0 };
  Object.keys(sheets).forEach((name, index) => {
    const sheet = createSheet({ name, sheetId: index + 1, state: 'visible', kind: 'worksheet', part: `/xl/worksheets/sheet${index + 1}.xml` });
    for (const address of Object.keys(sheets[name])) {
      const at = parseCellRef(address) as { row: number; col: number };
      place(sheet, at.row, at.col, makeCell(sheets[name][address]));
    }
    workbook.sheets.push(sheet);
  });
  return workbook;
}

/** Changes a cell the way an edit does: a new cell in a new row object. Returns the address for recalculate(). */
function edit(workbook: Workbook, sheet: number, address: string, spec: Spec | undefined): CellAddress {
  const at = parseCellRef(address) as { row: number; col: number };
  const model = workbook.sheets[sheet];
  const row = model.rows.get(at.row);
  const next: Row = row ? { ...row, cells: row.cells.slice() } : emptyRow();
  if (spec === undefined) delete next.cells[at.col];
  else next.cells[at.col] = makeCell(spec);
  model.rows.set(at.row, next);
  if (at.row >= model.rowCount) model.rowCount = at.row + 1;
  if (at.col >= model.colCount) model.colCount = at.col + 1;
  return { sheet, row: at.row, col: at.col };
}

function valueAt(workbook: Workbook, address: string, sheet = 0): Plain | undefined {
  const cell = parseCellRef(address) as { row: number; col: number };
  const found = getCell(workbook.sheets[sheet], cell.row, cell.col);
  return found ? found.value : undefined;
}

const addresses = (cells: CellAddress[]): string[] => cells.map((cell) => `${cell.sheet}!${cellName(cell.row, cell.col)}`).sort();

/** The value a formula gets in a cell of a sheet holding `cells`. */
function result(formula: string | FormulaSpec, at = 'H2', cells: Cells = {}, names: DefinedName[] = []): Plain | undefined {
  const workbook = book({ S: { ...cells, [at]: typeof formula === 'string' ? `=${formula}` : formula } }, names);
  new CalcEngine(workbook).recalculate();
  return valueAt(workbook, at);
}

function close(actual: unknown, expected: number, label: string): void {
  assert.ok(typeof actual === 'number' && Math.abs(actual - expected) <= 1e-9 * Math.max(1, Math.abs(expected)), `${label}: ${String(actual)} is not ${expected}`);
}

// ---------------------------------------------------------------------------
// 1. Implicit intersection

const COLUMNS: Cells = { A1: 10, A2: 20, A3: 30, B1: 1, B2: 2, B3: 3 };

test('a range meeting an operator in an ordinary formula stands for the cell in the formula row', () => {
  assert.equal(result('A1:A3*2', 'H2', COLUMNS), 40);
  assert.equal(result('A1:A3+B1:B3', 'H2', COLUMNS), 22);
  assert.equal(result('A1:A3&"x"', 'H2', COLUMNS), '20x');
  assert.equal(result('-A1:A3', 'H3', COLUMNS), -30);
  assert.equal(result('A1:A3%', 'H1', COLUMNS), 0.1);
  assert.equal(result('A1:A3>15', 'H1', COLUMNS), false);
  assert.equal(result('A:A*2', 'H3', COLUMNS), 60);
  assert.equal(result('A1:A3', 'H3', COLUMNS), 30);
});

test('a range in a row stands for the cell in the formula column', () => {
  const row: Cells = { A1: 1, B1: 2, C1: 3 };
  assert.equal(result('A1:C1*10', 'B5', row), 20);
  assert.equal(result('1:1*10', 'C5', row), 30);
  assert.equal(result('A1:C1*10', 'E5', row), '#VALUE!');
});

test('a range that does not cross the formula row or column is #VALUE!; one cell is itself', () => {
  assert.equal(result('A1:A3*2', 'H5', COLUMNS), '#VALUE!');
  assert.equal(result('A1:B3*2', 'H2', COLUMNS), '#VALUE!');
  assert.equal(result('A2*2', 'H5', COLUMNS), 40);
  assert.equal(result('A2:A2*2', 'H5', COLUMNS), 40);
});

test('a range given where a function takes one value is intersected as well', () => {
  assert.equal(result('ABS(A1:A3)', 'H2', COLUMNS), 20);
  assert.equal(result('IF(A1:A3>15,"y","n")', 'H2', COLUMNS), 'y');
  assert.equal(result('IF(A1:A3>15,"y","n")', 'H1', COLUMNS), 'n');
  assert.equal(result('ROUND(A1:A3/3,0)', 'H2', COLUMNS), 7);
  assert.equal(result('VLOOKUP(A1:A3,A1:B3,2,FALSE)', 'H2', COLUMNS), 2);
  assert.equal(result('LEN(A1:A3)', 'H3', COLUMNS), 2);
  assert.equal(result('LEFT(A1:A3,1)&"!"', 'H3', COLUMNS), '3!');
  assert.equal(result('NOT(A1:A3>15)', 'H1', COLUMNS), true);
  assert.equal(result('IFERROR(A1:A3/0,"e")', 'H1', COLUMNS), 'e');
  assert.equal(result('ABS(A1:A3)', 'H5', COLUMNS), '#VALUE!');
});

test('defined names are intersected the same way', () => {
  const names: DefinedName[] = [
    { name: 'Data', value: 'S!$A$1:$A$3', hidden: false },
    { name: 'Rates', value: 'S!$B$1:$B$3', hidden: false },
  ];
  assert.equal(result('Data*2', 'H2', COLUMNS, names), 40);
  assert.equal(result('Data*Rates', 'H3', COLUMNS, names), 90);
  assert.equal(result('ABS(Data)', 'H1', COLUMNS, names), 10);
  assert.equal(result('Data*2', 'H5', COLUMNS, names), '#VALUE!');
  assert.equal(result('SUM(Data)', 'H5', COLUMNS, names), 60);
  assert.equal(result('SUMPRODUCT(Data*Rates)', 'H5', COLUMNS, names), 140);
});

test('functions that take ranges still get the whole range', () => {
  assert.equal(result('SUM(A1:A3)', 'H2', COLUMNS), 60);
  assert.equal(result('SUM(A1:A3)', 'H9', COLUMNS), 60);
  assert.equal(result('COUNTIF(A1:A3,">15")', 'H9', COLUMNS), 2);
  assert.equal(result('SUMIF(A1:A3,">15",B1:B3)', 'H9', COLUMNS), 5);
  assert.equal(result('MAX(A1:A3,B1:B3)', 'H9', COLUMNS), 30);
  assert.equal(result('INDEX(A1:B3,3,2)', 'H9', COLUMNS), 3);
  assert.equal(result('ROWS(A1:A3)', 'H9', COLUMNS), 3);
  assert.equal(result('AND(A1:A3>5)', 'H2', COLUMNS), true);
});

test('array arguments are calculated element by element even in an ordinary formula', () => {
  assert.equal(result('SUMPRODUCT(A1:A3*B1:B3)', 'H2', COLUMNS), 140);
  assert.equal(result('SUMPRODUCT(A1:A3*B1:B3)', 'H9', COLUMNS), 140);
  assert.equal(result('SUMPRODUCT(A1:A3,B1:B3)', 'H9', COLUMNS), 140);
  assert.equal(result('SUMPRODUCT(--(A1:A3>15),B1:B3)', 'H9', COLUMNS), 5);
  assert.equal(result('SUMPRODUCT(ABS(-A1:A3))', 'H9', COLUMNS), 60);
  // The classic "last match" lookup.
  assert.equal(result('LOOKUP(2,1/(A1:A3>15),B1:B3)', 'H9', COLUMNS), 3);
  assert.equal(result('MATCH(60,A1:A3*2,0)', 'H9', COLUMNS), 3);
  assert.equal(result('INDEX(A1:A3*2,3)', 'H9', COLUMNS), 60);
  assert.equal(result('SUM(INDEX(A1:A3*B1:B3,0))', 'H9', COLUMNS), 140);
  assert.equal(result('VLOOKUP(40,A1:B3*2,2,FALSE)', 'H9', COLUMNS), 4);
  assert.equal(result('SUMPRODUCT(ROW(A1:A3))', 'H9', COLUMNS), 6);
  assert.equal(result('SUMPRODUCT(COUNTIF(A1:A3,B1:B3*10))', 'H9', COLUMNS), 3);
  assert.equal(result('SUMPRODUCT((A1:A3>15)*(B1:B3>2))', 'H9', COLUMNS), 1);
  assert.equal(result('INDEX(B1:B3,MATCH(1,(A1:A3>15)*(B1:B3>2),0))', 'H9', COLUMNS), 3);
  assert.equal(result('AGGREGATE(14,6,B1:B3/(A1:A3<25),1)', 'H9', COLUMNS), 2);
  assert.equal(result('LARGE(A1:A3*B1:B3,1)', 'H9', COLUMNS), 90);
  // The reverse lookup: a table put together from two columns.
  assert.equal(result('VLOOKUP(2,IF({1,0},B1:B3,A1:A3),2,FALSE)', 'H9', COLUMNS), 20);
});

test('SUM of a product of ranges: one row in an ordinary formula, every row in an array formula', () => {
  assert.equal(result('SUM(A1:A3*B1:B3)', 'H2', COLUMNS), 40);
  assert.equal(result('SUM(A1:A3*B1:B3)', 'H9', COLUMNS), '#VALUE!');
  assert.equal(result({ f: 'SUM(A1:A3*B1:B3)', type: 'array', ref: 'H2' }, 'H2', COLUMNS), 140);
  assert.equal(result({ f: 'SUM(A1:A3*B1:B3)', type: 'array', ref: 'H9' }, 'H9', COLUMNS), 140);
  assert.equal(result({ f: 'SUM(IF(A1:A3>15,B1:B3))', type: 'array', ref: 'H9' }, 'H9', COLUMNS), 5);
  assert.equal(result({ f: 'SUM(ABS(A1:A3))', type: 'array', ref: 'H9' }, 'H9', COLUMNS), 60);
});

test('a dynamic array formula (a cell with cell metadata) is calculated as an array formula', () => {
  assert.equal(result({ f: 'SUM(A1:A3*B1:B3)', cm: 1 }, 'H9', COLUMNS), 140);
  assert.equal(result({ f: 'A1:A3*2', cm: 1 }, 'H9', COLUMNS), 20);
});

test('an array formula over several cells still fills them', () => {
  const workbook = book({ S: { ...COLUMNS, D1: { f: 'A1:A3*B1:B3', type: 'array', ref: 'D1:D3' } } });
  new CalcEngine(workbook).recalculate();
  assert.deepEqual([valueAt(workbook, 'D1'), valueAt(workbook, 'D2'), valueAt(workbook, 'D3')], [10, 40, 90]);
});

test('array constants are arrays in an ordinary formula too', () => {
  assert.equal(result('SUM({1,2,3}*2)', 'H9'), 12);
  assert.equal(result('SUM(ABS({-1,2,-3}))', 'H9'), 6);
});

test('_xlfn.SINGLE is the implicit intersection operator', () => {
  assert.equal(result('_xlfn.SINGLE(A1:A3)', 'H2', COLUMNS), 20);
  assert.equal(result('_xlfn.SINGLE(A1:A3)', 'H7', COLUMNS), '#VALUE!');
});

test('the cells of a shared formula intersect on their own rows', () => {
  const workbook = book({
    S: {
      ...COLUMNS,
      H1: { f: 'A$1:A$3*2', type: 'shared', ref: 'H1:H3' },
      H2: { f: '', type: 'shared' },
      H3: { f: '', type: 'shared' },
    },
  });
  const cells = workbook.sheets[0].rows;
  for (const row of [0, 1, 2]) (cells.get(row) as Row).cells[7].formula!.si = 0;
  new CalcEngine(workbook).recalculate();
  assert.deepEqual([valueAt(workbook, 'H1'), valueAt(workbook, 'H2'), valueAt(workbook, 'H3')], [20, 40, 60]);
});

// ---------------------------------------------------------------------------
// 2. Numbers as text

test('a number joined to text is written out with up to 15 significant digits', () => {
  assert.equal(result('123456789012&""'), '123456789012');
  assert.equal(result('100000000000&""'), '100000000000');
  assert.equal(result('1/3&""'), '0.333333333333333');
  assert.equal(result('1234567.891234&""'), '1234567.891234');
  assert.equal(result('LEN(1/3)'), 17);
  assert.equal(result('0.1+0.2&""'), '0.3');
  assert.equal(result('1E+20&""'), '1E+20');
  assert.equal(result('1E-10&""'), '1E-10');
  assert.equal(result('123456789012345678&""'), '1.23456789012346E+17');
  assert.equal(result('-1/3&""'), '-0.333333333333333');
  assert.equal(result('0&""'), '0');
  assert.equal(result('-2.5&""'), '-2.5');
  assert.equal(result('123456789012345&""'), '123456789012345');
  assert.equal(result('0.001&""'), '0.001');
  assert.equal(result('2/3&""'), '0.666666666666667');
  assert.equal(result('1.5E+300&""'), '1.5E+300');
  assert.equal(result('CONCATENATE(1/8,"|",1E+15)'), '0.125|1E+15');
  assert.equal(result('LEFT(1234567.891234,9)'), '1234567.8');
  assert.equal(result('TEXTJOIN("-",TRUE,123456789012,1/4)'), '123456789012-0.25');
  // UNSURE: where Excel switches to scientific notation for small numbers with many digits.
  assert.equal(result('1/3000&""'), '0.000333333333333333');
});

// ---------------------------------------------------------------------------
// 3. Text as numbers: operators follow the same rule as function arguments

test('operators read dates, times, currency and thousands separators in text', () => {
  assert.equal(result('"1/1/2020"+0'), 43831);
  assert.equal(result('"1,000"+1'), 1001);
  assert.equal(result('"$5"+1'), 6);
  assert.equal(result('"5%"+0'), 0.05);
  close(result('"1:30"*24'), 1.5, '"1:30"*24');
  assert.equal(result('" 12 "+1'), 13);
  assert.equal(result('"abc"+1'), '#VALUE!');
  assert.equal(result('-"3"'), -3);
  assert.equal(result('"2020-03-01"-"2020-02-01"'), 29);
  assert.equal(result('""+1'), '#VALUE!');
  assert.equal(result('A1*2', 'H2', { A1: '1,5' }), '#VALUE!');
  // The same text through a function argument.
  assert.equal(result('ABS("1,000")'), 1000);
  assert.equal(result('SUM("1/1/2020",1)'), 43832);
});

// ---------------------------------------------------------------------------
// 4. SUBTOTAL and AGGREGATE leave out other subtotals

const SUBTOTALS: Cells = { E1: 1, E2: 2, E3: '=SUBTOTAL(9,E1:E2)', E4: 3, E5: '=SUBTOTAL(9,E4:E4)' };

test('SUBTOTAL ignores cells that hold a SUBTOTAL or AGGREGATE formula', () => {
  assert.equal(result('SUBTOTAL(9,E1:E5)', 'H1', SUBTOTALS), 6);
  assert.equal(result('SUBTOTAL(109,E1:E5)', 'H1', SUBTOTALS), 6);
  assert.equal(result('SUBTOTAL(2,E1:E5)', 'H1', SUBTOTALS), 3);
  assert.equal(result('SUBTOTAL(3,E1:E5)', 'H1', SUBTOTALS), 3);
  assert.equal(result('SUBTOTAL(1,E1:E5)', 'H1', SUBTOTALS), 2);
  assert.equal(result('SUBTOTAL(4,E1:E5)', 'H1', SUBTOTALS), 3);
  assert.equal(result('SUBTOTAL(9,E1:E5)', 'H1', { ...SUBTOTALS, E5: '=AGGREGATE(9,4,E4:E4)' }), 6);
  // SUM does not care.
  assert.equal(result('SUM(E1:E5)', 'H1', SUBTOTALS), 12);
  // Text that merely mentions the function is not a subtotal.
  assert.equal(result('SUBTOTAL(9,E1:E3)', 'H1', { E1: 1, E2: 2, E3: '=LEN("SUBTOTAL(")' }), 12);
});

test('AGGREGATE ignores nested subtotals for options 0 to 3 only', () => {
  assert.equal(result('AGGREGATE(9,0,E1:E5)', 'H1', SUBTOTALS), 6);
  assert.equal(result('AGGREGATE(9,3,E1:E5)', 'H1', SUBTOTALS), 6);
  assert.equal(result('AGGREGATE(9,4,E1:E5)', 'H1', SUBTOTALS), 12);
  assert.equal(result('AGGREGATE(9,6,E1:E5)', 'H1', SUBTOTALS), 12);
  assert.equal(result('AGGREGATE(14,0,E1:E5,1)', 'H1', SUBTOTALS), 3);
});

// ---------------------------------------------------------------------------
// 5. YEARFRAC, basis 0

test('YEARFRAC basis 0 keeps an end on the 31st when the start is the last day of February', () => {
  close(result('YEARFRAC(DATE(2019,2,28),DATE(2019,3,31),0)'), 31 / 360, 'Feb 28 to Mar 31');
  close(result('YEARFRAC(DATE(2019,2,28),DATE(2019,3,31))'), 31 / 360, 'Feb 28 to Mar 31, basis omitted');
  close(result('YEARFRAC(DATE(2020,2,29),DATE(2020,3,31),0)'), 31 / 360, 'leap Feb 29 to Mar 31');
  close(result('YEARFRAC(DATE(2019,2,28),DATE(2020,2,29),0)'), 1, 'Feb end to Feb end');
  close(result('YEARFRAC(DATE(2019,1,31),DATE(2019,3,31),0)'), 60 / 360, '31st to 31st');
  close(result('YEARFRAC(DATE(2019,1,30),DATE(2019,3,31),0)'), 60 / 360, '30th to 31st');
  close(result('YEARFRAC(DATE(2019,1,29),DATE(2019,3,31),0)'), 62 / 360, '29th to 31st');
  close(result('YEARFRAC(DATE(2019,1,31),DATE(2019,2,28),0)'), 28 / 360, '31st to Feb end');
  close(result('YEARFRAC(DATE(2020,2,28),DATE(2020,3,31),0)'), 33 / 360, 'Feb 28 of a leap year is not the last day');
});

test('DAYS360 (US) keeps its own rule: the last day of February as a start counts as a 30th', () => {
  // UNSURE: Excel's DAYS360 and YEARFRAC disagree here; this is the rule other engines implement for DAYS360.
  assert.equal(result('DAYS360(DATE(2019,2,28),DATE(2019,3,31))'), 30);
});

// ---------------------------------------------------------------------------
// 6. Formulas the engine cannot calculate keep the value of the file

test('a formula with an unknown function keeps its stored value', () => {
  const workbook = book({
    S: {
      A1: 5,
      B1: { f: 'GETPIVOTDATA("Sales",A1)', v: 1234 },
      B2: { f: 'CUBEVALUE("x",A1)+1', v: 'kept' },
      C1: '=B1*2',
      C2: '=B2&"!"',
    },
  });
  const engine = new CalcEngine(workbook);
  engine.recalculate();
  assert.equal(valueAt(workbook, 'B1'), 1234);
  assert.equal(valueAt(workbook, 'B2'), 'kept');
  // Formulas that read such a cell use the value it kept.
  assert.equal(valueAt(workbook, 'C1'), 2468);
  assert.equal(valueAt(workbook, 'C2'), 'kept!');
  assert.deepEqual(addresses(engine.unsupported), ['0!B1', '0!B2']);
  assert.deepEqual(engine.unparsed, []);
  // Changing what it reads does not touch it either.
  engine.recalculate([edit(workbook, 0, 'A1', 6)]);
  assert.equal(valueAt(workbook, 'B1'), 1234);
});

test('3D references, other workbooks, LET and LAMBDA keep their stored values', () => {
  const workbook = book({
    S: {
      A1: 1,
      B1: { f: 'SUM(S:T!A1)', v: 11 },
      B2: { f: '[1]Sheet1!A1*2', v: 22 },
      B3: { f: 'LET(x,A1,x+1)', v: 33 },
      B4: { f: 'LAMBDA(x,x+1)(A1)', v: 44 },
      B5: { f: "'[Book.xlsx]Sheet1'!A1", v: 55 },
      B6: { f: 'IF(A1>0,GETPIVOTDATA("a",A1),0)', v: 66 },
      B7: { f: 'Outside+1', v: 77 },
      // A name the workbook does not define may be a table: tables are not in the model.
      B8: { f: 'SUM(Table1)', v: 88 },
      B9: { f: 'Gone!A1+1', v: 99 },
      B10: { f: 'Known+1', v: 0 },
    },
    T: { A1: 2 },
  }, [{ name: 'Outside', value: '[1]Sheet1!$A$1', hidden: false }, { name: 'Known', value: 'S!$A$1', hidden: false }]);
  const engine = new CalcEngine(workbook);
  // Even when every value of the file may be replaced.
  engine.recalculate(undefined, { overwrite: true });
  const kept = ['B1', 'B2', 'B3', 'B5', 'B6', 'B7', 'B8', 'B9'].map((address) => valueAt(workbook, address));
  assert.deepEqual(kept, [11, 22, 33, 55, 66, 77, 88, 99]);
  assert.equal(valueAt(workbook, 'B10'), 2);
  // "LAMBDA(...)(...)" is not syntax the parser knows; it was already left alone.
  assert.equal(valueAt(workbook, 'B4'), 44);
  const listed = addresses(engine.unsupported.concat(engine.unparsed));
  assert.deepEqual(listed, ['0!B1', '0!B2', '0!B3', '0!B4', '0!B5', '0!B6', '0!B7', '0!B8', '0!B9']);
});

test('a formula just typed in (no value yet) shows the error instead of nothing', () => {
  const workbook = book({ S: { A1: 1, B1: '=NOSUCHFUNCTION(A1)', B2: '=SUM(S:T!A1)', B3: '=LET(x,1,x)', B4: '=NoSuchName+1', B5: '=Gone!A1' }, T: { A1: 2 } });
  const engine = new CalcEngine(workbook);
  engine.recalculate();
  assert.equal(valueAt(workbook, 'B1'), '#NAME?');
  assert.equal(valueAt(workbook, 'B2'), '#REF!');
  assert.equal(valueAt(workbook, 'B3'), '#NAME?');
  assert.equal(valueAt(workbook, 'B4'), '#NAME?');
  assert.equal(valueAt(workbook, 'B5'), '#REF!');
  assert.deepEqual(engine.unsupported, []);
  // The same when it is typed over a calculated cell.
  const typed = edit(workbook, 0, 'A1', '=NOSUCHFUNCTION(1)');
  engine.recalculate([typed]);
  assert.equal(valueAt(workbook, 'A1'), '#NAME?');
});

test('an unsupported formula is listed again after the graph is rebuilt, and dropped when it is replaced', () => {
  const workbook = book({ S: { B1: { f: 'GETPIVOTDATA("Sales",A1)', v: 1 } } });
  const engine = new CalcEngine(workbook);
  engine.recalculate();
  engine.invalidate();
  engine.recalculate();
  assert.deepEqual(addresses(engine.unsupported), ['0!B1']);
  engine.recalculate([edit(workbook, 0, 'B1', '=1+1')]);
  assert.deepEqual(engine.unsupported, []);
  assert.equal(valueAt(workbook, 'B1'), 2);
});

test('HYPERLINK shows its friendly name', () => {
  assert.equal(result('HYPERLINK("https://example.com","Site")'), 'Site');
  assert.equal(result('HYPERLINK("https://example.com")'), 'https://example.com');
});

// ---------------------------------------------------------------------------
// 7. Circular references do not depend on the order of calculation

test('every cell of a circular group is circular and 0, whatever is calculated first', () => {
  const cells: Cells = { A1: '=B1+C1', B1: '=A1', C1: '=B1+5', D1: '=C1+1', E1: 7, F1: '=E1*2' };
  const full = book({ S: cells });
  const engine = new CalcEngine(full);
  engine.recalculate();
  assert.deepEqual(['A1', 'B1', 'C1', 'D1', 'F1'].map((address) => valueAt(full, address)), [0, 0, 0, 1, 14]);
  assert.deepEqual(addresses(engine.circular), ['0!A1', '0!B1', '0!C1']);
  for (const start of ['A1', 'B1', 'C1']) {
    const at = parseCellRef(start) as { row: number; col: number };
    engine.recalculate([{ sheet: 0, row: at.row, col: at.col }]);
    assert.deepEqual(['A1', 'B1', 'C1', 'D1', 'F1'].map((address) => valueAt(full, address)), [0, 0, 0, 1, 14], `from ${start}`);
    assert.deepEqual(addresses(engine.circular), ['0!A1', '0!B1', '0!C1'], `from ${start}`);
  }
  // The same cells entered in another order.
  const reversed = book({ S: { F1: cells.F1, D1: cells.D1, C1: cells.C1, B1: cells.B1, A1: cells.A1, E1: 7 } });
  const other = new CalcEngine(reversed);
  other.recalculate();
  assert.deepEqual(['A1', 'B1', 'C1', 'D1', 'F1'].map((address) => valueAt(reversed, address)), [0, 0, 0, 1, 14]);
  assert.deepEqual(addresses(other.circular), ['0!A1', '0!B1', '0!C1']);
});

test('a cycle found after one of its formulas was calculated starts the pass over', () => {
  // A1 is tied to the cycle by a branch it never takes, so it is calculated (5) before A2 finds the cycle;
  // C1 has read that 5 by then.
  const cells: Cells = { A1: '=IF(FALSE,B2,5)', C1: '=A1+1', A2: '=B2+A1', B2: '=A2', C2: '=A2+A1+10' };
  const workbook = book({ S: cells });
  const engine = new CalcEngine(workbook);
  engine.recalculate();
  assert.deepEqual(['A1', 'C1', 'A2', 'B2', 'C2'].map((address) => valueAt(workbook, address)), [0, 1, 0, 0, 10]);
  assert.deepEqual(addresses(engine.circular), ['0!A1', '0!A2', '0!B2']);
  engine.recalculate([{ sheet: 0, row: 1, col: 1 }]);
  assert.deepEqual(['A1', 'C1', 'A2', 'B2', 'C2'].map((address) => valueAt(workbook, address)), [0, 1, 0, 0, 10]);
});

test('a cycle through cells chosen while calculating is found as well', () => {
  const workbook = book({ S: { A1: '=OFFSET(A1,1,0)+1', A2: '=INDIRECT("A1")+1', A3: '=A2+1', B1: '=OFFSET(B1,1,0)+1', B2: 4 } });
  const engine = new CalcEngine(workbook);
  engine.recalculate();
  assert.deepEqual(['A1', 'A2', 'A3', 'B1'].map((address) => valueAt(workbook, address)), [0, 0, 1, 5]);
  assert.deepEqual(addresses(engine.circular), ['0!A1', '0!A2']);
});

test('a formula that names its own cell in a branch that is not calculated is circular', () => {
  const workbook = book({ S: { H20: '=IF(FALSE,H20,1)', H21: '=H20+1', J1: '=SUM(J1:J3)', K5: '=ROW(K5)', L5: '=COLUMNS(L5:M5)', M1: '=INDEX(M:M,2)', M2: 8, N1: '=OFFSET(N1,1,0)', N2: 9, P1: '=INDEX(P:P,1)' } });
  const engine = new CalcEngine(workbook);
  engine.recalculate();
  // UNSURE: a range a function picks cells from does not make the formula circular unless it picks its own cell.
  assert.equal(valueAt(workbook, 'M1'), 8);
  assert.equal(valueAt(workbook, 'N1'), 9);
  assert.equal(valueAt(workbook, 'P1'), 0);
  assert.equal(valueAt(workbook, 'H20'), 0);
  assert.equal(valueAt(workbook, 'H21'), 1);
  assert.equal(valueAt(workbook, 'J1'), 0);
  // Asking where a cell is does not read it.
  assert.equal(valueAt(workbook, 'K5'), 5);
  assert.equal(valueAt(workbook, 'L5'), 2);
  assert.deepEqual(addresses(engine.circular), ['0!H20', '0!J1', '0!P1']);
});

/** A small deterministic generator, so that a failure can be reproduced from its seed. */
function generator(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return state / 4294967296;
  };
}

function shuffled<T>(items: T[], random: () => number): T[] {
  const out = items.slice();
  for (let k = out.length - 1; k > 0; k--) {
    const j = Math.floor(random() * (k + 1));
    const held = out[k];
    out[k] = out[j];
    out[j] = held;
  }
  return out;
}

test('random graphs with cycles: a full recalculation equals an incremental one from every cell', () => {
  for (let seed = 1; seed <= 120; seed++) {
    const random = generator(seed);
    const count = 4 + Math.floor(random() * 9);
    const cellsOf = (k: number) => cellName(Math.floor(k / 3), k % 3);
    const cells: Cells = {};
    for (let k = 0; k < count; k++) {
      const kind = random();
      if (kind < 0.2) {
        cells[cellsOf(k)] = 1 + Math.floor(random() * 9);
        continue;
      }
      const pick = () => cellsOf(Math.floor(random() * count));
      let formula: string;
      if (kind < 0.55) formula = `${pick()}+${pick()}+${1 + Math.floor(random() * 5)}`;
      else if (kind < 0.7) formula = `SUM(${pick()}:${pick()})+1`;
      else if (kind < 0.85) formula = `IF(${pick()}>3,${pick()},${pick()}+2)`;
      else formula = `IF(FALSE,${pick()},${pick()}*2)`;
      cells[cellsOf(k)] = `=${formula}`;
    }
    const keys = Object.keys(cells);
    const values = (workbook: Workbook) => keys.map((key) => valueAt(workbook, key));
    const label = `seed ${seed}: ${JSON.stringify(cells)}`;

    const reference = book({ S: cells });
    const first = new CalcEngine(reference);
    first.recalculate();
    const expected = values(reference);
    const circular = addresses(first.circular);

    // The same cells put into the sheet in another order.
    const order: Cells = {};
    for (const key of shuffled(keys, random)) order[key] = cells[key];
    const mixed = book({ S: order });
    const second = new CalcEngine(mixed);
    second.recalculate();
    assert.deepEqual(values(mixed), expected, `other order, ${label}`);
    assert.deepEqual(addresses(second.circular), circular, `other order, ${label}`);

    // Calculating again changes nothing, from wherever it starts.
    first.recalculate();
    assert.deepEqual(values(reference), expected, `second full run, ${label}`);
    for (const key of keys) {
      const at = parseCellRef(key) as { row: number; col: number };
      first.recalculate([{ sheet: 0, row: at.row, col: at.col }]);
      assert.deepEqual(values(reference), expected, `from ${key}, ${label}`);
      for (const address of addresses(first.circular)) assert.ok(circular.indexOf(address) >= 0, `${address} circular from ${key} only, ${label}`);
    }
  }
});

// ---------------------------------------------------------------------------
// 11. Arguments left empty

test('an argument left empty counts as 0, FALSE or "", not as the default', () => {
  const table: Cells = { A1: 1, A2: 3, A3: 5, B1: 'one', B2: 'three', B3: 'five' };
  // Exact match: 4 is not there. With the default (approximate) it would find 3.
  assert.equal(result('VLOOKUP(4,A1:B3,2,)', 'H9', table), '#N/A');
  assert.equal(result('VLOOKUP(4,A1:B3,2)', 'H9', table), 'three');
  assert.equal(result('VLOOKUP(3,A1:B3,2,)', 'H9', table), 'three');
  assert.equal(result('HLOOKUP(4,A1:C1,1,)', 'H9', { A1: 1, B1: 3, C1: 5 }), '#N/A');
  assert.equal(result('MATCH(4,A1:A3,)', 'H9', table), '#N/A');
  assert.equal(result('MATCH(4,A1:A3)', 'H9', table), 2);
  // A reference to an empty cell does the same.
  assert.equal(result('VLOOKUP(4,A1:B3,2,Z1)', 'H9', table), '#N/A');
  assert.equal(result('MATCH(4,A1:A3,Z1)', 'H9', table), '#N/A');
  assert.equal(result('ROUND(1.5,)'), 2);
  assert.equal(result('ROUND(1.5,0)'), 2);
  assert.equal(result('IF(TRUE,,2)'), 0);
  assert.equal(result('IF(FALSE,1,)'), 0);
  assert.equal(result('IF(FALSE,1)'), false);
  assert.equal(result('LEFT("abc",)'), '');
  assert.equal(result('LEFT("abc")'), 'a');
  assert.equal(result('TRUNC(2.7,)'), 2);
  assert.equal(result('SUM(1,,2)'), 3);
  assert.equal(result('PMT(0,10,1000,,)'), -100);
  // UNSURE: the same rule applied to the other classic functions with a default that is not 0.
  assert.equal(result('FIXED(1234.567,,TRUE)'), '1235');
  assert.equal(result('FIXED(1234.567)'), '1,234.57');
  assert.equal(result('DOLLAR(1.5,)'), '$2');
  assert.equal(result('WEEKDAY(DATE(2024,3,15),)'), '#NUM!');
  assert.equal(result('WEEKDAY(DATE(2024,3,15))'), 6);
  assert.equal(result('ROUND(IRR({-100,60,60},),6)'), 0.130662);
  assert.equal(result('CHOOSE(1,,5)'), 0);
  assert.equal(result('INDEX(A1:B3,,2)', 'H9', table), '#VALUE!');
  assert.equal(result('SUM(INDEX(A1:A3,,1))', 'H9', table), 9);
  // The functions of the dynamic array generation take an empty argument as "not given".
  assert.equal(result('XLOOKUP(4,A1:A3,B1:B3,,0)', 'H9', table), '#N/A');
  assert.equal(result('XLOOKUP(4,A1:A3,B1:B3,,-1)', 'H9', table), 'three');
  assert.equal(result('OFFSET(A1,1,0,,)', 'H9', table), 3);
});

// ---------------------------------------------------------------------------
// 12. The rest

test('operator precedence', () => {
  assert.equal(result('2*3^2'), 18);
  assert.equal(result('1+2*3^2'), 19);
  assert.equal(result('-2^2'), 4);
  assert.equal(result('2^3^2'), 64);
  assert.equal(result('2+3&4+5'), '59');
  assert.equal(result('1+2=3'), true);
  assert.equal(result('10-4-3'), 3);
  assert.equal(result('100/10/5'), 2);
  assert.equal(result('2^-1'), 0.5);
  assert.equal(result('50%*2'), 1);
  assert.equal(result('-3%'), -0.03);
});

test('a range built while calculating makes the formula volatile', () => {
  const workbook = book({ S: { A1: 1, A2: 2, A3: 3, B1: 2, C1: '=SUM(A1:INDEX(A1:A3,B1))', C2: '=SUM(A1:A2)' } });
  const engine = new CalcEngine(workbook);
  engine.recalculate();
  assert.equal(valueAt(workbook, 'C1'), 3);
  engine.recalculate([edit(workbook, 0, 'B1', 3)]);
  assert.equal(valueAt(workbook, 'C1'), 6);
  // Volatile: calculated although nothing it names changed. Both formulas call SUM; only C1 runs again.
  let calls = 0;
  const sum = FUNCTIONS.get('SUM') as FunctionSpec;
  const counting = new Map(FUNCTIONS);
  counting.set('SUM', { ...sum, call: (args, context) => (calls++, sum.call(args, context)) });
  const second = new CalcEngine(workbook, { functions: counting });
  second.recalculate();
  assert.equal(calls, 2);
  calls = 0;
  second.recalculate([edit(workbook, 0, 'Z9', 1)]);
  assert.equal(calls, 1, 'only the formula with the built range is calculated again');
});

test('looking up an empty cell looks for 0', () => {
  // UNSURE: Excel reads an empty lookup value as 0; it finds a cell holding 0, never an empty cell.
  const table: Cells = { A1: 5, A2: 0, A3: 7, B1: 'five', B2: 'zero', B3: 'seven' };
  assert.equal(result('VLOOKUP(Z1,A1:B3,2,FALSE)', 'H9', table), 'zero');
  assert.equal(result('MATCH(Z1,A1:A3,0)', 'H9', table), 2);
  assert.equal(result('VLOOKUP(Z1,A1:B3,2,FALSE)', 'H9', { A1: 5, A3: 7, B1: 'five', B3: 'seven' }), '#N/A');
});

test('FIXED keeps the minus sign of a negative number that rounds to zero', () => {
  assert.equal(result('FIXED(-0.004,2)'), '-0.00');
  assert.equal(result('FIXED(0.004,2)'), '0.00');
  assert.equal(result('FIXED(-1234.567,1)'), '-1,234.6');
  assert.equal(result('FIXED(0,2)'), '0.00');
});
