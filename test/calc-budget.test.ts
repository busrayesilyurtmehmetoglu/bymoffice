// Tests for what keeps a recalculation within bounds: text that cannot make a conversion
// slow, the limit on the text one formula builds, the budget of a recalculation, and the
// index that answers exact lookups and "equals" criteria.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine, type CalcEngineOptions, type CellAddress } from '../src/calc/engine';
import { FUNCTIONS } from '../src/calc/functions/index';
import { compilePattern, patternMatches, patternSearch } from '../src/calc/functions/helpers';
import { MAX_FORMULA_TEXT, numberToText, parseNumberText, refToArray, textToNumber, toNumber, type CalcContext, type FunctionSpec } from '../src/calc/values';
import { cellName, createSheet, emptyStyleSheet, getCell, parseCellRef, type Cell, type CellFormula, type Row, type Sheet, type Workbook } from '../src/model/sheet/types';

type Plain = number | string | boolean | null;
interface FormulaSpec {
  f: string;
  v?: Plain;
  type?: CellFormula['type'];
  ref?: string;
}
type Spec = Plain | FormulaSpec;
type Cells = Record<string, Spec>;

function makeCell(spec: Spec): Cell {
  if (typeof spec === 'string' && spec.charAt(0) === '=') spec = { f: spec.slice(1) };
  if (spec !== null && typeof spec === 'object') {
    const value = spec.v === undefined ? null : spec.v;
    const formula: CellFormula = { text: spec.f, type: spec.type || 'normal' };
    if (spec.ref !== undefined) formula.ref = spec.ref;
    return { kind: value === null ? 'blank' : typeof value === 'number' ? 'number' : typeof value === 'boolean' ? 'boolean' : 'string', value, style: 0, formula };
  }
  return { kind: spec === null ? 'blank' : typeof spec === 'number' ? 'number' : typeof spec === 'boolean' ? 'boolean' : 'string', value: spec, style: 0 };
}

const emptyRow = (): Row => ({ customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] });

function put(sheet: Sheet, row: number, col: number, spec: Spec): void {
  let target = sheet.rows.get(row);
  if (!target) sheet.rows.set(row, (target = emptyRow()));
  target.cells[col] = makeCell(spec);
  if (row >= sheet.rowCount) sheet.rowCount = row + 1;
  if (col >= sheet.colCount) sheet.colCount = col + 1;
}

function book(cells: Cells = {}): Workbook {
  const workbook: Workbook = { sheets: [], styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 };
  const sheet = createSheet({ name: 'S', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  for (const address of Object.keys(cells)) {
    const at = parseCellRef(address) as { row: number; col: number };
    put(sheet, at.row, at.col, cells[address]);
  }
  workbook.sheets.push(sheet);
  return workbook;
}

/** Changes a cell the way an edit does: a new cell in a new row object. */
function edit(workbook: Workbook, address: string, spec: Spec): CellAddress {
  const at = parseCellRef(address) as { row: number; col: number };
  const model = workbook.sheets[0];
  const row = model.rows.get(at.row);
  const next: Row = row ? { ...row, cells: row.cells.slice() } : emptyRow();
  next.cells[at.col] = makeCell(spec);
  model.rows.set(at.row, next);
  if (at.row >= model.rowCount) model.rowCount = at.row + 1;
  if (at.col >= model.colCount) model.colCount = at.col + 1;
  return { sheet: 0, row: at.row, col: at.col };
}

function valueAt(workbook: Workbook, address: string): Plain | undefined {
  const at = parseCellRef(address) as { row: number; col: number };
  const cell = getCell(workbook.sheets[0], at.row, at.col);
  return cell ? cell.value : undefined;
}

const NO_LIMIT: CalcEngineOptions = { budget: { milliseconds: 0 } };

function result(formula: string | FormulaSpec, cells: Cells = {}, at = 'H1'): Plain | undefined {
  const workbook = book({ ...cells, [at]: typeof formula === 'string' ? `=${formula}` : formula });
  new CalcEngine(workbook, NO_LIMIT).recalculate();
  return valueAt(workbook, at);
}

function repeat(text: string, count: number): string {
  let out = '';
  for (let k = 0; k < count; k++) out += text;
  return out;
}

/** Milliseconds `run` takes: the best of three, so that a busy machine does not fail the test. */
function timed(run: () => void): number {
  let best = Infinity;
  for (let k = 0; k < 3; k++) {
    const start = process.hrtime.bigint();
    run();
    best = Math.min(best, Number(process.hrtime.bigint() - start) / 1e6);
  }
  return best;
}

// ---------------------------------------------------------------------------
// Text as numbers

test('plain numbers in text', () => {
  const cases: [string, number | undefined][] = [
    ['12', 12], [' 12 ', 12], ['-3.5', -3.5], ['+7', 7], ['.5', 0.5], ['5.', 5], ['1e3', 1000], ['1E-2', 0.01], ['1.5e+2', 150],
    ['50%', 0.5], ['50 %', 0.5], [' 2.5 % ', 0.025], ['\t8\n', 8],
    ['', undefined], [' ', undefined], ['.', undefined], ['-', undefined], ['1e', undefined], ['1e+', undefined], ['e5', undefined],
    ['1 2', undefined], ['12x', undefined], ['%', undefined], ['5%%', undefined], ['--1', undefined], ['1.2.3', undefined],
    ['0x10', undefined], ['Infinity', undefined], ['NaN', undefined], ['1e999', undefined],
  ];
  for (const [text, expected] of cases) assert.equal(textToNumber(text), expected, JSON.stringify(text));
});

test('text that reads as a number, wherever a number is needed', () => {
  assert.equal(parseNumberText('1,000'), 1000);
  assert.equal(parseNumberText('$5'), 5);
  assert.equal(parseNumberText('1/1/2020'), 43831);
  assert.equal(parseNumberText('abc'), undefined);
  assert.equal(parseNumberText('=1+1'), undefined);
  assert.equal(parseNumberText("'12"), undefined);
  assert.equal(parseNumberText('TRUE'), undefined);
  assert.equal(toNumber('1,000'), 1000);
  assert.deepEqual(toNumber('abc'), { error: '#VALUE!' });
  assert.equal(toNumber(null), 0);
  assert.equal(toNumber(true), 1);
});

test('numbers as text', () => {
  const cases: [number, string][] = [
    [0, '0'], [-0, '0'], [1, '1'], [-1, '-1'], [0.5, '0.5'], [123456789012, '123456789012'], [1e14, '100000000000000'],
    [123456789012345, '123456789012345'], [1e15, '1E+15'], [1234567890123456, '1.23456789012346E+15'], [1e20, '1E+20'],
    [1.5e300, '1.5E+300'], [1 / 3, '0.333333333333333'], [2 / 3, '0.666666666666667'], [0.1 + 0.2, '0.3'], [1e-9, '0.000000001'],
    [1e-10, '1E-10'], [-2.5e-12, '-2.5E-12'], [1234567.891234, '1234567.891234'], [0.1, '0.1'], [100, '100'], [99999999999999.9, '99999999999999.9'],
    [999999999999999.9, '1E+15'], [1e-7, '0.0000001'], [5e-324, '4.94065645841247E-324'],
  ];
  for (const [value, expected] of cases) assert.equal(numberToText(value), expected, String(value));
  // What comes out reads back as the same number, to 15 digits.
  let seed = 7;
  for (let k = 0; k < 2000; k++) {
    seed = (Math.imul(seed, 1103515245) + 12345) >>> 0;
    const value = (seed / 4294967296 - 0.5) * Math.pow(10, (k % 40) - 20);
    const text = numberToText(value);
    const back = Number(text);
    assert.ok(Math.abs(back - value) <= Math.abs(value) * 1e-14, `${value} -> ${text}`);
    assert.ok(text.length <= 22, text);
  }
});

// ---------------------------------------------------------------------------
// Hostile text is not slow

test('a long run of digits followed by a letter is rejected at once', () => {
  const hostile = repeat('1', 32767) + 'x';
  const took = timed(() => {
    for (let k = 0; k < 50; k++) assert.equal(textToNumber(hostile), undefined);
  });
  assert.ok(took < 250, `50 conversions took ${took} ms`);
  const variants = [repeat('1', 32767), repeat(' ', 32767) + 'x', repeat('1', 16000) + '.' + repeat('1', 16000) + 'e', repeat('1.', 16000), repeat('%', 32767), '1' + repeat(' ', 32000) + '%' + repeat(' ', 700) + 'x', repeat('1e', 16000), repeat('-', 32767)];
  const each = timed(() => {
    for (const text of variants) {
      textToNumber(text);
      parseNumberText(text);
    }
  });
  assert.ok(each < 250, `the variants took ${each} ms`);
});

test('formulas over hostile text finish quickly', () => {
  const cases: [string, Plain][] = [
    ['IFERROR(SUMPRODUCT(--(REPT("1",32000)&"x"&ROW(A1:A200))),-1)', -1],
    ['IFERROR(VALUE(REPT("1",32767)&"x"),-1)', -1],
    ['IFERROR(VALUE(REPT("1 ",16000)),-1)', -1],
    ['IFERROR(NUMBERVALUE(REPT("1",1000)&"x"),-1)', -1],
    ['IFERROR(NUMBERVALUE(REPT("1",32767)),-1)', -1],
    ['IFERROR(DATEVALUE(REPT("1/",16000)),-1)', -1],
    ['IFERROR(TIMEVALUE(REPT("1:",16000)),-1)', -1],
    ['IFERROR(REPT("9",32767)+1,-1)', -1],
    ['IFERROR(REPT(" ",32767)&"x"+1,-1)', -1],
    ['COUNTIF(A1:A3,REPT("1",32767)&"x")', 0],
    ['COUNTIF(A1:A3,">"&REPT("1",32767))', 0],
    ['IFERROR(SEARCH(REPT("a*",8000)&"b",REPT("a",32767)),-1)', -1],
    ['IFERROR(SEARCH(REPT("?",300)&"b",REPT("a",32767)),-1)', -1],
    ['COUNTIF(A1:A3,REPT("*a",8000)&"b")', 0],
    ['IFERROR(MATCH(REPT("a?",8000)&"b",A1:A3,0),-1)', -1],
    ['LEN(SUBSTITUTE(REPT("ab",16000),"b",""))', 16000],
    ['LEN(TRIM(REPT(" a",16000)))', 31999],
    ['LEN(TEXT(1,REPT("0",250)))', 250],
    ['IFERROR(LEN(TEXT(REPT("1",32767),"0")),-1)', 32767],
  ];
  const cells: Cells = { A1: repeat('a', 32767), A2: repeat('1', 32767), A3: repeat('a ', 16000) };
  for (const [formula, expected] of cases) {
    let actual: Plain | undefined;
    const took = timed(() => {
      actual = result(formula, cells);
    });
    assert.equal(actual, expected, formula);
    assert.ok(took < 1500, `${formula} took ${took} ms`);
  }
});

test('wildcard matching costs at most text length times pattern length', () => {
  const text = repeat('a', 32767);
  const took = timed(() => {
    assert.equal(patternMatches(compilePattern(repeat('*a', 16000) + '*b'), text), false);
    assert.equal(patternMatches(compilePattern(repeat('a*', 16000)), text), true);
    assert.equal(patternSearch(compilePattern(repeat('a*', 8000) + 'b'), text, 0), -1);
  });
  assert.ok(took < 1500, `took ${took} ms`);
});

// ---------------------------------------------------------------------------
// The text one formula may build

test('a formula that would build gigabytes of text is #VALUE!, not a crash', () => {
  const formulas: FormulaSpec[] = [
    { f: 'SUM(LEN(SUBSTITUTE(ROW(A1:A100000)&REPT("x",32000),"x","y")))', type: 'array', ref: 'H1' },
    { f: 'SUM(LEN(REPT("x",32000+0*ROW(A1:A100000))))', type: 'array', ref: 'H1' },
    { f: 'SUMPRODUCT(LEN(REPT("x",32000)&ROW(A1:A100000)))' },
    { f: 'SUM(LEN(IF(ROW(A1:A100000)>0,REPT("x",32000)&ROW(A1:A100000))))', type: 'array', ref: 'H1' },
  ];
  for (const formula of formulas) {
    let actual: Plain | undefined;
    const took = timed(() => {
      actual = result(formula);
    });
    assert.equal(actual, '#VALUE!', formula.f);
    // A guard against a hang, not a measurement: about 2 s on an idle machine, several times that under load.
    assert.ok(took < 30000, `${formula.f} took ${took} ms`);
  }
  // Well inside the limit it works.
  assert.equal(result({ f: 'SUM(LEN(REPT("x",1000)&ROW(A1:A1000)))', type: 'array', ref: 'H1' }), 1000 * 1000 + 9 + 90 * 2 + 900 * 3 + 4);
  assert.ok(MAX_FORMULA_TEXT >= 1000 * 1004);
  // The count starts again for every formula.
  const workbook = book({ H1: { f: 'SUM(LEN(REPT("x",30000)&ROW(A1:A1000)))', type: 'array', ref: 'H1' }, H2: { f: 'SUM(LEN(REPT("y",30000)&ROW(A1:A1000)))', type: 'array', ref: 'H2' }, H3: '=H1+H2' });
  new CalcEngine(workbook, NO_LIMIT).recalculate();
  assert.equal(typeof valueAt(workbook, 'H1'), 'number');
  assert.equal(valueAt(workbook, 'H3'), 2 * (valueAt(workbook, 'H1') as number));
});

// ---------------------------------------------------------------------------
// References that are not rectangles

test('refToArray never throws', () => {
  const context: CalcContext = {
    sheet: 0, row: 0, col: 0, date1904: false,
    cellValue: () => null,
    forEachCell: () => undefined,
    usedExtent: () => ({ rows: 1048576, cols: 16384 }),
    sheetIndex: () => 0, sheetName: () => 'S', sheetCount: () => 1, now: () => 0, random: () => 0,
  };
  const one = (top: number, left: number, bottom: number, right: number) => refToArray({ ref: true, sheet: 0, range: { top, left, bottom, right } }, context);
  assert.deepEqual(one(5, 0, 2, 0), { rows: 1, cols: 1, data: [{ error: '#REF!' }] });
  assert.deepEqual(one(0, 5, 0, 2), { rows: 1, cols: 1, data: [{ error: '#REF!' }] });
  assert.deepEqual(one(-1, 0, 2, 0), { rows: 1, cols: 1, data: [{ error: '#REF!' }] });
  assert.deepEqual(one(NaN, 0, 2, 0), { rows: 1, cols: 1, data: [{ error: '#REF!' }] });
  assert.deepEqual(one(0, 0, 1048575, 16383), { rows: 1, cols: 1, data: [{ error: '#NUM!' }] });
  assert.deepEqual(one(0, 0, 1.5, 0), { rows: 1, cols: 1, data: [{ error: '#REF!' }] });
  assert.equal(one(0, 0, 1, 1).data.length, 4);
  // A formula over the whole sheet as an array is an error value, not an exception.
  const workbook = book({ A1: 1, XFD1048576: 2, H1: { f: 'SUM(J:XFD*2)', type: 'array', ref: 'H1' } });
  new CalcEngine(workbook, NO_LIMIT).recalculate();
  assert.equal(valueAt(workbook, 'H1'), '#NUM!');
});

// ---------------------------------------------------------------------------
// The budget of a recalculation

/** 200 numbers in column A and 300 formulas in column B that each read all of them. */
function costly(): { cells: Cells; expected: (row: number, first: number) => number } {
  const cells: Cells = {};
  for (let r = 1; r <= 200; r++) cells[`A${r}`] = r;
  for (let r = 1; r <= 300; r++) cells[`B${r}`] = `=SUM(A$1:A$200)+ROW()`;
  return { cells, expected: (row, first) => 20100 - 1 + first + row };
}

test('a recalculation stops when its work budget is used up and goes on in the next call', () => {
  const { cells, expected } = costly();
  const workbook = book(cells);
  const engine = new CalcEngine(workbook, { budget: { work: 5000, milliseconds: 0 } });
  engine.recalculate();
  assert.equal(engine.incomplete, true);
  let calculated = 0;
  for (let r = 1; r <= 300; r++) {
    const value = valueAt(workbook, `B${r}`);
    // Never a half-calculated value: the right one, or the one it had.
    if (value !== null) {
      assert.equal(value, expected(r, 1), `B${r}`);
      calculated++;
    }
  }
  assert.ok(calculated > 0 && calculated < 300, `${calculated} formulas in the first call`);
  let calls = 1;
  while (engine.incomplete) {
    const changed = engine.recalculate([]);
    assert.ok(changed.length > 0, 'every call makes progress');
    assert.ok(++calls < 200);
  }
  assert.ok(calls > 2);
  for (let r = 1; r <= 300; r++) assert.equal(valueAt(workbook, `B${r}`), expected(r, 1), `B${r}`);
  assert.deepEqual(engine.overBudget, []);
  // Complete: another call has nothing to do.
  assert.deepEqual(engine.recalculate([]), []);
});

test('a change made while a recalculation is incomplete is taken into account', () => {
  const { cells, expected } = costly();
  const workbook = book(cells);
  const engine = new CalcEngine(workbook, { budget: { work: 5000, milliseconds: 0 } });
  engine.recalculate();
  assert.equal(engine.incomplete, true);
  let changed: CellAddress[] = [edit(workbook, 'A1', 1001)];
  for (let calls = 0; calls < 200; calls++) {
    engine.recalculate(changed);
    changed = [];
    if (!engine.incomplete) break;
  }
  assert.equal(engine.incomplete, false);
  for (let r = 1; r <= 300; r++) assert.equal(valueAt(workbook, `B${r}`), expected(r, 1001), `B${r}`);
});

test('formulas that wait on one another are finished in order across calls', () => {
  const cells: Cells = {};
  for (let r = 1; r <= 100; r++) cells[`A${r}`] = 1;
  // Each formula reads the one below it: calculating the first needs all of them.
  for (let r = 1; r <= 150; r++) cells[`B${r}`] = r === 150 ? '=SUM(A$1:A$100)' : `=B${r + 1}+SUM(A$1:A$100)`;
  const workbook = book(cells);
  const engine = new CalcEngine(workbook, { budget: { work: 3000, milliseconds: 0 } });
  engine.recalculate();
  assert.equal(engine.incomplete, true);
  for (let calls = 0; engine.incomplete; calls++) {
    for (let r = 1; r <= 150; r++) {
      const value = valueAt(workbook, `B${r}`);
      if (value !== null) assert.equal(value, (151 - r) * 100, `B${r} after ${calls} more calls`);
    }
    engine.recalculate([]);
    assert.ok(calls < 500);
  }
  for (let r = 1; r <= 150; r++) assert.equal(valueAt(workbook, `B${r}`), (151 - r) * 100);
});

test('the time budget is read from the clock that was given', () => {
  const { cells, expected } = costly();
  const workbook = book(cells);
  let now = 1000;
  let reads = 0;
  const engine = new CalcEngine(workbook, {
    budget: { milliseconds: 10 },
    clock: () => {
      reads++;
      return (now += 1);
    },
  });
  engine.recalculate();
  assert.equal(engine.incomplete, true, 'ten ticks of the clock are not enough');
  assert.ok(reads >= 10 && reads < 40, `the clock was read ${reads} times`);
  let calls = 1;
  while (engine.incomplete) {
    engine.recalculate([]);
    assert.ok(++calls < 200);
  }
  for (let r = 1; r <= 300; r++) assert.equal(valueAt(workbook, `B${r}`), expected(r, 1));
  // A clock that stands still never stops anything.
  const still = book(cells);
  const patient = new CalcEngine(still, { budget: { milliseconds: 10 }, clock: () => 5 });
  patient.recalculate();
  assert.equal(patient.incomplete, false);
  // No limit at all.
  const free = book(cells);
  const unlimited = new CalcEngine(free, { budget: { milliseconds: 0 }, clock: () => (now += 1000) });
  unlimited.recalculate();
  assert.equal(unlimited.incomplete, false);
});

test('a formula that cannot be calculated within the hard budget keeps its value and is not tried forever', () => {
  const cells: Cells = { D1: { f: 'SUMPRODUCT(COUNTIF(A1:A300,">"&ROW(A1:A300)))', v: 123 }, D2: '=D1+1', E1: '=A1+1', E2: '=A2*2' };
  for (let r = 1; r <= 300; r++) cells[`A${r}`] = r;
  const workbook = book(cells);
  const engine = new CalcEngine(workbook, { budget: { work: 20000, hardWork: 40000, milliseconds: 0 } });
  engine.recalculate(undefined, { overwrite: true });
  let calls = 1;
  while (engine.incomplete) {
    engine.recalculate([]);
    assert.ok(++calls < 10, 'the expensive formula is given up');
  }
  assert.deepEqual(engine.overBudget, [{ sheet: 0, row: 0, col: 3 }]);
  assert.deepEqual(engine.stale, [{ sheet: 0, row: 0, col: 3 }, { sheet: 0, row: 1, col: 3 }]);
  assert.equal(valueAt(workbook, 'D1'), 123);
  assert.equal(valueAt(workbook, 'D2'), 124);
  assert.equal(valueAt(workbook, 'E1'), 2);
  assert.equal(valueAt(workbook, 'E2'), 4);
  // A change of what it reads tries it again, once, and gives it up again.
  engine.recalculate([edit(workbook, 'A1', 5)]);
  calls = 1;
  while (engine.incomplete) {
    engine.recalculate([]);
    assert.ok(++calls < 10, 'the expensive formula is given up again');
  }
  assert.equal(valueAt(workbook, 'D1'), 123);
  assert.equal(valueAt(workbook, 'E1'), 6);
  assert.deepEqual(engine.overBudget, [{ sheet: 0, row: 0, col: 3 }]);
  // A change of something else does not.
  assert.deepEqual(engine.recalculate([edit(workbook, 'G1', 5)]), []);
  assert.equal(engine.incomplete, false);
  // A new formula in the cell is a new start.
  engine.recalculate([edit(workbook, 'D1', '=COUNT(A1:A300)')]);
  assert.equal(valueAt(workbook, 'D1'), 300);
  assert.equal(valueAt(workbook, 'D2'), 301);
  assert.deepEqual(engine.overBudget, []);
  assert.deepEqual(engine.stale, []);
  // With the default hard budget (ten times the budget, a million units at least) the same formula is calculated,
  // in a call that is allowed to take longer because nothing else was done in it.
  const roomy = book(cells);
  const patient = new CalcEngine(roomy, { budget: { work: 20000, milliseconds: 0 } });
  patient.recalculate(undefined, { overwrite: true });
  calls = 1;
  while (patient.incomplete) {
    patient.recalculate([]);
    assert.ok(++calls < 10);
  }
  assert.equal(valueAt(roomy, 'D1'), 44850);
  assert.equal(valueAt(roomy, 'D2'), 44851);
  assert.deepEqual(patient.overBudget, []);
});

test('a preview that runs out of budget is #VALUE!', () => {
  const cells: Cells = {};
  for (let r = 1; r <= 300; r++) cells[`A${r}`] = r;
  const workbook = book(cells);
  const engine = new CalcEngine(workbook, { budget: { work: 20000, milliseconds: 0 } });
  assert.deepEqual(engine.evaluateText('SUMPRODUCT(COUNTIF(A1:A300,">"&ROW(A1:A300)))', { sheet: 0, row: 0, col: 5 }), { error: '#VALUE!' });
  assert.equal(engine.evaluateText('SUM(A1:A300)', { sheet: 0, row: 0, col: 5 }), 45150);
});

test('work that reads no cells counts as well', () => {
  const workbook = book({ A1: 1 });
  const tight = new CalcEngine(workbook, { budget: { work: 5000, milliseconds: 0 } });
  const at = { sheet: 0, row: 0, col: 5 };
  // 100,000 elements through an operator, and through a function applied to each element.
  assert.deepEqual(tight.evaluateText('SUMPRODUCT(ROW(A1:A100000)*2)', at), { error: '#VALUE!' });
  assert.deepEqual(tight.evaluateText('SUMPRODUCT(ABS(ROW(A1:A100000)))', at), { error: '#VALUE!' });
  assert.deepEqual(tight.evaluateText('SUMPRODUCT(--(-ROW(A1:A100000)))', at), { error: '#VALUE!' });
  assert.equal(tight.evaluateText('SUMPRODUCT(ROW(A1:A1000)*2)', at), 1001000);
  const free = new CalcEngine(workbook, NO_LIMIT);
  assert.equal(free.evaluateText('SUMPRODUCT(ABS(ROW(A1:A100000)))', at), 5000050000);
});

test('the default budget stops a recalculation that takes too long', () => {
  // 20,000 criteria that each scan 20,000 cells: far more than five seconds of work.
  const cells: Cells = { H1: { f: 'SUMPRODUCT(COUNTIF(A:A,">"&ROW(A1:A20000)))', v: 7 } };
  const workbook = book(cells);
  const sheet = workbook.sheets[0];
  for (let r = 0; r < 20000; r++) put(sheet, r, 0, r + 1);
  let now = 0;
  // A clock that jumps a second at every look: the test does not have to wait.
  const engine = new CalcEngine(workbook, { clock: () => (now += 1000) });
  const start = Date.now();
  engine.recalculate();
  // It had the call to itself and went on to the hard limit (ten times the budget): it is given up, and that is said.
  assert.equal(engine.incomplete, false);
  assert.deepEqual(engine.overBudget, [{ sheet: 0, row: 0, col: 7 }]);
  assert.deepEqual(engine.stale, [{ sheet: 0, row: 0, col: 7 }]);
  assert.equal(valueAt(workbook, 'H1'), 7);
  assert.ok(now >= 50000 && now < 80000, `the clock went to ${now}`);
  assert.ok(Date.now() - start < 3000);
});

// ---------------------------------------------------------------------------
// The index for exact lookups and "equals" criteria

/** A library whose functions never see the memo of the context: every search is a scan. */
function scanning(): Map<string, FunctionSpec> {
  const library = new Map<string, FunctionSpec>();
  FUNCTIONS.forEach((spec, name) => {
    library.set(name, {
      ...spec,
      call: (args, context) => {
        const bare = Object.create(context) as { memo?: unknown };
        bare.memo = undefined;
        return spec.call(args, bare as CalcContext);
      },
    });
  });
  return library;
}

test('searching through the index gives what scanning gives', () => {
  let seed = 99;
  const random = () => {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    return seed / 4294967296;
  };
  const keys: Plain[] = [1, 2, 3, 2.5, 0, -1, '1', '2', ' 3', 'abc', 'ABC', 'Abc', 'x*', 'a?c', '', true, false, 'TRUE', '1e0', '0', 'İstanbul', 'istanbul', repeat('long', 100), repeat('LONG', 100)];
  for (let round = 0; round < 12; round++) {
    const cells: Cells = {};
    const rows = 40 + Math.floor(random() * 60);
    for (let r = 1; r <= rows; r++) {
      if (random() < 0.85) cells[`A${r}`] = keys[Math.floor(random() * keys.length)];
      if (random() < 0.05) cells[`A${r}`] = '=1/0';
      if (random() < 0.05) cells[`A${r}`] = '=1+1';
      cells[`B${r}`] = random() < 0.9 ? r : 'text';
      cells[`C${r}`] = random() < 0.5 ? 'p' : 'q';
    }
    const formulas: string[] = [];
    const literal = (key: Plain) => (typeof key === 'string' ? `"${key}"` : typeof key === 'boolean' ? (key ? 'TRUE' : 'FALSE') : String(key));
    for (const key of keys.concat([7, 'zzz', null])) {
      const k = key === null ? 'Z1' : literal(key);
      formulas.push(`VLOOKUP(${k},A1:B${rows},2,FALSE)`, `MATCH(${k},A1:A${rows},0)`, `MATCH(${k},A:A,0)`, `XLOOKUP(${k},A1:A${rows},B1:B${rows},"none",0,-1)`, `XLOOKUP(${k},A1:A${rows},B1:B${rows})`);
      formulas.push(`COUNTIF(A1:A${rows},${k})`, `COUNTIF(A:A,${k})`, `SUMIF(A1:A${rows},${k},B1:B${rows})`, `COUNTIFS(A1:A${rows},${k},C1:C${rows},"p")`, `COUNTIFS(C1:C${rows},"p",A1:A${rows},${k})`);
      formulas.push(`AVERAGEIF(A1:A${rows},${k},B1:B${rows})`, `SUMIFS(B1:B${rows},A1:A${rows},${k},C1:C${rows},"q")`, `MAXIFS(B1:B${rows},A1:A${rows},${k})`);
      if (typeof key === 'string' && key !== '') formulas.push(`COUNTIF(A1:A${rows},"=${key}")`, `COUNTIF(A1:A${rows},"<>${key}")`);
    }
    // Twice, so that the second half of the formulas certainly meets ranges that were searched before.
    const all = formulas.concat(formulas);
    all.forEach((formula, index) => {
      cells[cellName(index, 5)] = `=${formula}`;
    });
    const indexed = book(cells);
    const engine = new CalcEngine(indexed, NO_LIMIT);
    engine.recalculate();
    const scanned = book(cells);
    new CalcEngine(scanned, { ...NO_LIMIT, functions: scanning() }).recalculate();
    all.forEach((formula, index) => {
      const name = cellName(index, 5);
      assert.deepEqual(valueAt(indexed, name), valueAt(scanned, name), `round ${round}: ${formula}`);
    });
    // And again after an edit of the table: nothing of the old index is left.
    const address = edit(indexed, 'A3', 'abc');
    edit(scanned, 'A3', 'abc');
    engine.recalculate([address]);
    new CalcEngine(scanned, { ...NO_LIMIT, functions: scanning() }).recalculate();
    all.forEach((formula, index) => {
      const name = cellName(index, 5);
      assert.deepEqual(valueAt(indexed, name), valueAt(scanned, name), `round ${round} after an edit: ${formula}`);
    });
  }
});

test('exact lookups: kinds do not mix, text ignores case, the first match wins', () => {
  const table: Cells = { A1: 5, A2: '5', A3: 'abc', A4: 'ABC', A5: true, A6: 5, B1: 'n1', B2: 't', B3: 'lower', B4: 'upper', B5: 'yes', B6: 'n2' };
  for (const repeatCount of [1, 3]) {
    const cells: Cells = { ...table };
    const formulas = ['VLOOKUP(5,A1:B6,2,FALSE)', 'VLOOKUP("5",A1:B6,2,FALSE)', 'VLOOKUP("ABC",A1:B6,2,FALSE)', 'VLOOKUP(TRUE,A1:B6,2,FALSE)', 'VLOOKUP(1,A1:B6,2,FALSE)', 'XLOOKUP(5,A1:A6,B1:B6,,0,-1)', 'MATCH("abc",A1:A6,0)', 'VLOOKUP("a*",A1:B6,2,FALSE)', 'COUNTIF(A1:A6,5)', 'COUNTIF(A1:A6,"abc")'];
    const expected = ['n1', 't', 'lower', 'yes', '#N/A', 'n2', 3, 'lower', 3, 2];
    for (let k = 0; k < repeatCount; k++) formulas.forEach((formula, index) => (cells[cellName(index + k * 20, 7)] = `=${formula}`));
    const workbook = book(cells);
    new CalcEngine(workbook, NO_LIMIT).recalculate();
    for (let k = 0; k < repeatCount; k++) formulas.forEach((formula, index) => assert.equal(valueAt(workbook, cellName(index + k * 20, 7)), expected[index], `${formula} (${k})`));
  }
});

test('a lookup table whose keys are formulas, some of them circular', () => {
  const cells: Cells = { A1: '=1+1', A2: '=A1+1', A3: '=A3+1', A4: '=A2*2', B1: 'two', B2: 'three', B3: 'loop', B4: 'six' };
  for (let k = 0; k < 6; k++) cells[`D${k + 1}`] = `=VLOOKUP(${[2, 3, 0, 6, 2, 0][k]},A1:B4,2,FALSE)`;
  const workbook = book(cells);
  new CalcEngine(workbook, NO_LIMIT).recalculate();
  assert.deepEqual([1, 2, 3, 4, 5, 6].map((k) => valueAt(workbook, `D${k}`)), ['two', 'three', 'loop', 'six', 'two', 'loop']);
});

function lookupBook(rows: number, lookups: number): Workbook {
  const workbook = book();
  const sheet = workbook.sheets[0];
  for (let r = 0; r < rows; r++) {
    put(sheet, r, 0, r % 2 === 0 ? r * 7 : `key${r}`);
    put(sheet, r, 1, r);
  }
  for (let k = 0; k < lookups; k++) {
    const r = Math.floor(((k + 0.5) * rows) / lookups);
    put(sheet, k, 3, `=VLOOKUP(${r % 2 === 0 ? r * 7 : `"KEY${r}"`},A$1:B$${rows},2,FALSE)`);
    put(sheet, k, 4, r);
  }
  return workbook;
}

test('performance: 1,000 exact lookups in a table of 100,000 rows', () => {
  const workbook = lookupBook(100000, 1000);
  const engine = new CalcEngine(workbook, NO_LIMIT);
  const took = timed(() => {
    engine.recalculate();
  });
  for (let k = 0; k < 1000; k++) assert.equal(valueAt(workbook, cellName(k, 3)), valueAt(workbook, cellName(k, 4)), `row ${k}`);
  console.log(`# 1,000 VLOOKUP(...,FALSE) x 100,000 rows: ${took.toFixed(1)} ms`);
  assert.ok(took < 2000, `took ${took} ms`);
});

test('performance: editing one cell of a table that 4,000 lookups read', () => {
  const workbook = lookupBook(100000, 4000);
  const engine = new CalcEngine(workbook, NO_LIMIT);
  engine.recalculate();
  let turn = 0;
  const took = timed(() => {
    engine.recalculate([edit(workbook, 'B50001', 1000000 + ++turn)]);
  });
  assert.equal(engine.incomplete, false);
  for (let k = 0; k < 4000; k++) {
    const expected = valueAt(workbook, cellName(k, 4));
    assert.equal(valueAt(workbook, cellName(k, 3)), expected === 50000 ? 1000000 + turn : expected, `row ${k}`);
  }
  console.log(`# one edit, 4,000 dependent VLOOKUP(...,FALSE) x 100,000 rows: ${took.toFixed(1)} ms`);
  assert.ok(took < 2000, `took ${took} ms`);
});

test('performance: 20,000 "equals" criteria over a column of 20,000 cells', () => {
  const workbook = book({ H1: '=SUMPRODUCT(COUNTIF(A:A,ROW(A1:A20000)))', H2: '=SUMPRODUCT(SUMIF(A:A,ROW(A1:A20000),B:B))' });
  const sheet = workbook.sheets[0];
  for (let r = 0; r < 20000; r++) {
    put(sheet, r, 0, (r % 5000) + 1);
    put(sheet, r, 1, 1);
  }
  const engine = new CalcEngine(workbook, NO_LIMIT);
  const took = timed(() => {
    engine.recalculate();
  });
  // Every value from 1 to 5,000 is there four times; the criteria 5,001 to 20,000 match nothing.
  assert.equal(valueAt(workbook, 'H1'), 20000);
  assert.equal(valueAt(workbook, 'H2'), 20000);
  console.log(`# SUMPRODUCT(COUNTIF(A:A,ROW(A1:A20000))) and the same with SUMIF: ${took.toFixed(1)} ms`);
  assert.ok(took < 2000, `took ${took} ms`);
});

test('a single lookup scans without building an index', () => {
  let builds = 0;
  const library = new Map(FUNCTIONS);
  const lookup = FUNCTIONS.get('VLOOKUP') as FunctionSpec;
  library.set('VLOOKUP', {
    ...lookup,
    call: (args, context) => {
      const counting = Object.create(context) as CalcContext;
      counting.memo = <T>(key: string, build: () => T): T => {
        return (context.memo as NonNullable<CalcContext['memo']>)(key, () => {
          if (key.lastIndexOf('index:', 0) === 0) builds++;
          return build();
        });
      };
      return lookup.call(args, counting);
    },
  });
  const cells: Cells = { A1: 1, A2: 2, A3: 3, B1: 'a', B2: 'b', B3: 'c', D1: '=VLOOKUP(2,A1:B3,2,FALSE)' };
  const once = book(cells);
  new CalcEngine(once, { ...NO_LIMIT, functions: library }).recalculate();
  assert.equal(valueAt(once, 'D1'), 'b');
  assert.equal(builds, 0);
  const thrice = book({ ...cells, D2: '=VLOOKUP(3,A1:B3,2,FALSE)', D3: '=VLOOKUP(1,A1:B3,2,FALSE)' });
  new CalcEngine(thrice, { ...NO_LIMIT, functions: library }).recalculate();
  assert.deepEqual(['D1', 'D2', 'D3'].map((name) => valueAt(thrice, name)), ['b', 'c', 'a']);
  assert.equal(builds, 1);
});
