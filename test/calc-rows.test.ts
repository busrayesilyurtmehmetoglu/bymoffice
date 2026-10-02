import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine, type CalcEngineOptions, type CellAddress, type RangeValues } from '../src/calc/engine';
import { FUNCTIONS } from '../src/calc/functions/index';
import { isError, type CalcError, type FunctionSpec, type Scalar } from '../src/calc/values';
import { DataValidations, MAX_LIST_ITEMS } from '../src/format/validation';
import {
  cellName,
  createSheet,
  emptyStyleSheet,
  getCell,
  parseCellRef,
  type AutoFilter,
  type Cell,
  type CellFormula,
  type DataValidation,
  type DefinedName,
  type Row,
  type Sheet,
  type Workbook,
} from '../src/model/sheet/types';
import { setAutoFilter, setFilterColumn } from '../src/ops/filter';
import { History } from '../src/ops/ops';

// Hidden rows (SUBTOTAL, AGGREGATE, rowsChanged), all values of a range (evaluateRange),
// circular references known before anything is calculated (cycles) and defined names that
// change (namesChanged).

// ---------------------------------------------------------------------------
// Building workbooks by hand

type Plain = number | string | boolean | null;
interface FormulaSpec {
  f: string;
  /** Value the file would have stored for the formula. */
  v?: Plain;
  error?: boolean;
}
type Spec = Plain | FormulaSpec;

function kindOf(value: Plain, error = false): Cell['kind'] {
  if (error) return 'error';
  return value === null ? 'blank' : typeof value === 'number' ? 'number' : typeof value === 'boolean' ? 'boolean' : 'string';
}

/** A string starting with "=" is a formula that was never calculated; one starting with "#" an error value. */
function makeCell(spec: Spec): Cell {
  if (typeof spec === 'string' && spec.charAt(0) === '=') spec = { f: spec.slice(1) };
  if (spec !== null && typeof spec === 'object') {
    const value = spec.v === undefined ? null : spec.v;
    const formula: CellFormula = { text: spec.f, type: 'normal' };
    return { kind: kindOf(value, spec.error), value, style: 0, formula };
  }
  if (typeof spec === 'string' && spec.charAt(0) === '#') return { kind: 'error', value: spec, style: 0 };
  return { kind: kindOf(spec), value: spec, style: 0 };
}

const emptyRow = (): Row => ({ customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] });

function book(sheets: Record<string, Record<string, Spec>>, definedNames: DefinedName[] = []): Workbook {
  const workbook: Workbook = { sheets: [], styles: emptyStyleSheet(), sharedStrings: [], definedNames, date1904: false, activeSheet: 0 };
  Object.keys(sheets).forEach((name, index) => {
    const sheet = createSheet({ name, sheetId: index + 1, state: 'visible', kind: 'worksheet', part: `/xl/worksheets/sheet${index + 1}.xml` });
    for (const address of Object.keys(sheets[name])) {
      const at = parseCellRef(address) as { row: number; col: number };
      let target = sheet.rows.get(at.row);
      if (!target) sheet.rows.set(at.row, (target = emptyRow()));
      target.cells[at.col] = makeCell(sheets[name][address]);
      if (at.row >= sheet.rowCount) sheet.rowCount = at.row + 1;
      if (at.col >= sheet.colCount) sheet.colCount = at.col + 1;
    }
    workbook.sheets.push(sheet);
  });
  return workbook;
}

/** Changes a cell the way an edit does: a new cell in a new row object. */
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

/** Hides or shows rows (1-based, as in A1) the way an edit does: new row objects. */
function hide(sheet: Sheet, rows: number[], hidden = true): void {
  for (const number of rows) {
    const row = sheet.rows.get(number - 1) || emptyRow();
    sheet.rows.set(number - 1, { ...row, hidden });
  }
}

/** A filter over a range with a criterion on its first column, or without any. */
function filterOn(range: string, active = true): AutoFilter {
  const [from, to] = range.split(':').map((name) => parseCellRef(name) as { row: number; col: number });
  return { range: { top: from.row, left: from.col, bottom: to.row, right: to.col }, columns: active ? [{ colId: 0, kind: 'custom', conditions: [] } as unknown as AutoFilter['columns'][number]] : [] };
}

const valueAt = (workbook: Workbook, address: string, sheet = 0) => {
  const cell = parseCellRef(address) as { row: number; col: number };
  return getCell(workbook.sheets[sheet], cell.row, cell.col)?.value;
};

function names(workbook: Workbook, result: { sheet: number; cells: { row: number; col: number }[] }[]): string[] {
  const out: string[] = [];
  for (const entry of result) for (const cell of entry.cells) out.push(`${workbook.sheets[entry.sheet].name}!${cellName(cell.row, cell.col)}`);
  return out.sort();
}

const cells = (list: CellAddress[]) => list.map((cell) => `${cell.sheet}!${cellName(cell.row, cell.col)}`).sort();

const NO_LIMIT: CalcEngineOptions = { budget: { milliseconds: 0 } };
const A1: CellAddress = { sheet: 0, row: 0, col: 0 };
const OUTSIDE: CellAddress = { sheet: 0, row: 20, col: 20 };

/** The real library with calls of some functions counted. */
function counting(...counted: string[]): { functions: Map<string, FunctionSpec>; calls: Record<string, number>; reset(): void } {
  const functions = new Map(FUNCTIONS);
  const calls: Record<string, number> = {};
  for (const name of counted) {
    const spec = FUNCTIONS.get(name) as FunctionSpec;
    calls[name] = 0;
    functions.set(name, {
      ...spec,
      call: (args, context) => {
        calls[name]++;
        return spec.call(args, context);
      },
    });
  }
  return {
    functions,
    calls,
    reset() {
      for (const name of counted) calls[name] = 0;
    },
  };
}

function code(value: Scalar | RangeValues | CalcError | undefined): Scalar | string | undefined {
  return isError(value) ? value.error : (value as Scalar);
}

// ---------------------------------------------------------------------------
// SUBTOTAL

/** A header and the numbers 1, 2, 4, 8, 16 under it. */
function list(): Workbook {
  return book({ S: { A1: 'n', A2: 1, A3: 2, A4: 4, A5: 8, A6: 16 } });
}

test('SUBTOTAL with every row shown counts every row', () => {
  const engine = new CalcEngine(list(), NO_LIMIT);
  assert.equal(engine.evaluateText('SUBTOTAL(9,A2:A6)', OUTSIDE), 31);
  assert.equal(engine.evaluateText('SUBTOTAL(109,A2:A6)', OUTSIDE), 31);
});

test('SUBTOTAL 9 counts a row hidden by hand, 109 does not', () => {
  const workbook = list();
  hide(workbook.sheets[0], [3]);
  const engine = new CalcEngine(workbook, NO_LIMIT);
  assert.equal(engine.evaluateText('SUBTOTAL(9,A2:A6)', OUTSIDE), 31);
  assert.equal(engine.evaluateText('SUBTOTAL(109,A2:A6)', OUTSIDE), 29);
  assert.equal(engine.evaluateText('SUM(A2:A6)', OUTSIDE), 31, 'SUM does not care');
  assert.equal(engine.evaluateText('COUNT(A2:A6)', OUTSIDE), 5);
});

test('SUBTOTAL 9 and 109 both leave out a row hidden by a filter', () => {
  const workbook = list();
  workbook.sheets[0].autoFilter = filterOn('A1:A6');
  hide(workbook.sheets[0], [3, 6]);
  const engine = new CalcEngine(workbook, NO_LIMIT);
  assert.equal(engine.evaluateText('SUBTOTAL(9,A2:A6)', OUTSIDE), 13);
  assert.equal(engine.evaluateText('SUBTOTAL(109,A2:A6)', OUTSIDE), 13);
  assert.equal(engine.evaluateText('SUM(A2:A6)', OUTSIDE), 31);
});

test('a hidden row counts as hidden by hand outside the filter and under a filter without a criterion', () => {
  // UNSURE: Excel knows which of the two hid a row; the model has one flag. A row hidden by hand
  // inside a list that is filtered is left out by Excel as well, which is what this rule gives.
  const idle = list();
  idle.sheets[0].autoFilter = filterOn('A1:A6', false);
  hide(idle.sheets[0], [3]);
  const first = new CalcEngine(idle, NO_LIMIT);
  assert.equal(first.evaluateText('SUBTOTAL(9,A2:A6)', OUTSIDE), 31);
  assert.equal(first.evaluateText('SUBTOTAL(109,A2:A6)', OUTSIDE), 29);

  const partly = list();
  partly.sheets[0].autoFilter = filterOn('A1:A4');
  hide(partly.sheets[0], [3, 6]);
  const second = new CalcEngine(partly, NO_LIMIT);
  assert.equal(second.evaluateText('SUBTOTAL(9,A2:A6)', OUTSIDE), 29, 'row 6 is under the filter range: hidden by hand');
  assert.equal(second.evaluateText('SUBTOTAL(109,A2:A6)', OUTSIDE), 13);
  // The header row is not a data row of the filter.
  hide(partly.sheets[0], [1]);
  assert.equal(second.evaluateText('SUBTOTAL(3,A1:A6)', OUTSIDE), 5);
  assert.equal(second.evaluateText('SUBTOTAL(103,A1:A6)', OUTSIDE), 3);
});

test('every function number of SUBTOTAL leaves hidden rows out', () => {
  const workbook = list();
  hide(workbook.sheets[0], [6]);
  const engine = new CalcEngine(workbook, NO_LIMIT);
  const expected: [number, number, number][] = [
    // number, with the hidden row, without it
    [1, 6.2, 3.75],
    [2, 5, 4],
    [3, 5, 4],
    [4, 16, 8],
    [5, 1, 1],
    [6, 1024, 64],
    [9, 31, 15],
  ];
  for (const [kind, all, shown] of expected) {
    assert.equal(engine.evaluateText(`SUBTOTAL(${kind},A2:A6)`, OUTSIDE), all, `SUBTOTAL ${kind}`);
    assert.equal(engine.evaluateText(`SUBTOTAL(${kind + 100},A2:A6)`, OUTSIDE), shown, `SUBTOTAL ${kind + 100}`);
  }
  assert.equal(code(engine.evaluateText('SUBTOTAL(12,A2:A6)', OUTSIDE)), '#VALUE!');
  assert.equal(code(engine.evaluateText('SUBTOTAL(112,A2:A6)', OUTSIDE)), '#VALUE!');
});

test('nested subtotals are passed over, hidden or not', () => {
  const workbook = book({
    S: { A1: 'n', A2: 1, A3: 2, A4: { f: 'SUBTOTAL(9,A2:A3)', v: 3 }, A5: 8, A6: 16, A7: { f: 'SUBTOTAL(109,A5:A6)', v: 24 }, A8: '=SUBTOTAL(9,A2:A7)', A9: '=SUBTOTAL(109,A2:A7)', A10: '=SUM(A2:A7)' },
  });
  const engine = new CalcEngine(workbook, NO_LIMIT);
  engine.recalculate();
  assert.equal(valueAt(workbook, 'A8'), 27);
  assert.equal(valueAt(workbook, 'A9'), 27);
  assert.equal(valueAt(workbook, 'A10'), 54);

  hide(workbook.sheets[0], [3, 6]);
  engine.rowsChanged(0);
  engine.recalculate([]);
  assert.equal(valueAt(workbook, 'A4'), 3, 'function number 9 counts the row hidden by hand');
  assert.equal(valueAt(workbook, 'A7'), 8);
  assert.equal(valueAt(workbook, 'A8'), 27);
  assert.equal(valueAt(workbook, 'A9'), 9);
  assert.equal(valueAt(workbook, 'A10'), 38, 'SUM reads the subtotals as they are now');
});

// ---------------------------------------------------------------------------
// AGGREGATE

test('AGGREGATE options 0 to 7 choose among hidden rows, errors and nested subtotals', () => {
  // Row 3 is hidden by hand, row 4 holds an error in column A, row 5 a subtotal.
  const workbook = book({
    S: {
      A2: 1,
      A3: 2,
      A4: '#DIV/0!',
      A5: { f: 'SUBTOTAL(9,A2:A3)', v: 3 },
      A6: 8,
      B2: 1,
      B3: 2,
      B5: { f: 'AGGREGATE(9,4,B2:B3)', v: 3 },
      B6: 8,
    },
  });
  hide(workbook.sheets[0], [3]);
  const engine = new CalcEngine(workbook, NO_LIMIT);
  const withError = ['#DIV/0!', '#DIV/0!', 11, 9, '#DIV/0!', '#DIV/0!', 14, 12];
  const clean = [11, 9, 11, 9, 14, 12, 14, 12];
  for (let option = 0; option <= 7; option++) {
    assert.equal(code(engine.evaluateText(`AGGREGATE(9,${option},A2:A6)`, OUTSIDE)), withError[option], `option ${option} with an error`);
    assert.equal(code(engine.evaluateText(`AGGREGATE(9,${option},B2:B6)`, OUTSIDE)), clean[option], `option ${option}`);
  }
  assert.equal(code(engine.evaluateText('AGGREGATE(9,8,B2:B6)', OUTSIDE)), '#VALUE!');
  // Counting and the forms with a second argument follow the same rules.
  assert.equal(engine.evaluateText('AGGREGATE(2,4,B2:B6)', OUTSIDE), 4);
  assert.equal(engine.evaluateText('AGGREGATE(2,5,B2:B6)', OUTSIDE), 3);
  assert.equal(engine.evaluateText('AGGREGATE(3,1,B2:B6)', OUTSIDE), 2);
  assert.equal(engine.evaluateText('AGGREGATE(14,4,B2:B6,2)', OUTSIDE), 3);
  assert.equal(engine.evaluateText('AGGREGATE(15,4,B2:B6,2)', OUTSIDE), 2);
  assert.equal(engine.evaluateText('AGGREGATE(15,5,B2:B6,2)', OUTSIDE), 3);
  assert.equal(engine.evaluateText('AGGREGATE(15,7,A2:A6,2)', OUTSIDE), 3);
  assert.equal(engine.evaluateText('AGGREGATE(15,3,A2:A6,2)', OUTSIDE), 8);
});

test('AGGREGATE counts rows a filter hides unless the option names hidden rows', () => {
  // UNSURE: Excel may leave filtered rows out with every option, as SUBTOTAL 1 to 11 does.
  const workbook = list();
  workbook.sheets[0].autoFilter = filterOn('A1:A6');
  hide(workbook.sheets[0], [3]);
  const engine = new CalcEngine(workbook, NO_LIMIT);
  assert.equal(engine.evaluateText('AGGREGATE(9,4,A2:A6)', OUTSIDE), 31);
  assert.equal(engine.evaluateText('AGGREGATE(9,5,A2:A6)', OUTSIDE), 29);
});

// ---------------------------------------------------------------------------
// rowsChanged

/** Ten numbers, a subtotal, a sum, a formula reading the subtotal; a second sheet with subtotals of its own and of the first. */
function subtotals(stored?: number): Workbook {
  const first: Record<string, Spec> = { B1: stored === undefined ? '=SUBTOTAL(109,A1:A10)' : { f: 'SUBTOTAL(109,A1:A10)', v: stored }, B2: '=SUM(A1:A10)', B3: '=B1*2' };
  for (let r = 1; r <= 10; r++) first[`A${r}`] = r;
  return book({ S: first, T: { A1: 5, A2: 6, C1: '=SUBTOTAL(109,A1:A2)', D1: '=AGGREGATE(9,5,S!A1:A10)', E1: '=SUM(A1:A2)' } });
}

test('rowsChanged() recalculates only the formulas that leave hidden rows out and read that sheet', () => {
  const workbook = subtotals();
  const counted = counting('SUBTOTAL', 'AGGREGATE', 'SUM');
  const engine = new CalcEngine(workbook, { ...NO_LIMIT, functions: counted.functions });
  engine.recalculate();
  assert.equal(valueAt(workbook, 'B1'), 55);
  assert.deepEqual(counted.calls, { SUBTOTAL: 2, AGGREGATE: 1, SUM: 2 });

  counted.reset();
  hide(workbook.sheets[0], [10]);
  engine.rowsChanged(0);
  assert.deepEqual(counted.calls, { SUBTOTAL: 0, AGGREGATE: 0, SUM: 0 }, 'nothing is calculated before recalculate()');
  const changed = engine.recalculate([]);
  assert.deepEqual(names(workbook, changed), ['S!B1', 'S!B3', 'T!D1']);
  assert.deepEqual(counted.calls, { SUBTOTAL: 1, AGGREGATE: 1, SUM: 0 }, 'the subtotal of the other sheet and the sums are left alone');
  assert.equal(valueAt(workbook, 'B1'), 45);
  assert.equal(valueAt(workbook, 'B2'), 55);
  assert.equal(valueAt(workbook, 'B3'), 90);
  assert.equal(valueAt(workbook, 'D1', 1), 45);
  assert.equal(engine.incomplete, false);

  // Once told, the engine does not calculate them again.
  counted.reset();
  assert.deepEqual(engine.recalculate([]), []);
  assert.deepEqual(counted.calls, { SUBTOTAL: 0, AGGREGATE: 0, SUM: 0 });

  // Showing the row again (an undo) brings the old values back.
  hide(workbook.sheets[0], [10], false);
  engine.rowsChanged(0);
  assert.deepEqual(names(workbook, engine.recalculate([])), ['S!B1', 'S!B3', 'T!D1']);
  assert.equal(valueAt(workbook, 'B1'), 55);
  assert.equal(valueAt(workbook, 'B3'), 110);

  // The other sheet, and every sheet at once.
  counted.reset();
  hide(workbook.sheets[1], [2]);
  engine.rowsChanged(1);
  assert.deepEqual(names(workbook, engine.recalculate([])), ['T!C1']);
  assert.equal(valueAt(workbook, 'C1', 1), 5);
  assert.deepEqual(counted.calls, { SUBTOTAL: 1, AGGREGATE: 0, SUM: 0 });
  counted.reset();
  engine.rowsChanged();
  assert.deepEqual(engine.recalculate([]), [], 'nothing changed in fact');
  assert.deepEqual(counted.calls, { SUBTOTAL: 2, AGGREGATE: 1, SUM: 0 });
  // A sheet that does not exist is ignored.
  engine.rowsChanged(7);
  engine.rowsChanged(-1);
  engine.rowsChanged(NaN);
  counted.reset();
  engine.recalculate([]);
  assert.deepEqual(counted.calls, { SUBTOTAL: 0, AGGREGATE: 0, SUM: 0 });
});

test('rows hidden by a real filter change SUBTOTAL 9, and undoing the filter changes it back', () => {
  const workbook = book({ S: { A1: 'fruit', B1: 'n', A2: 'apple', B2: 1, A3: 'pear', B3: 2, A4: 'apple', B4: 4, A5: 'plum', B5: 8, D1: '=SUBTOTAL(9,B2:B5)', D2: '=SUBTOTAL(109,B2:B5)', D3: '=SUM(B2:B5)' } });
  const engine = new CalcEngine(workbook, NO_LIMIT);
  engine.recalculate();
  const history = new History<undefined>(workbook);
  history.apply(setAutoFilter(workbook, 0, { top: 0, left: 0, bottom: 4, right: 1 }), undefined, undefined);
  engine.namesChanged();
  engine.rowsChanged(0);
  assert.deepEqual(engine.recalculate([]), [], 'a filter without a criterion hides nothing');

  history.apply(setFilterColumn(workbook, 0, 0, { kind: 'values', values: ['apple'] }), undefined, undefined);
  assert.equal(workbook.sheets[0].rows.get(2)?.hidden, true);
  engine.rowsChanged(0);
  assert.deepEqual(names(workbook, engine.recalculate([])), ['S!D1', 'S!D2']);
  assert.equal(valueAt(workbook, 'D1'), 5);
  assert.equal(valueAt(workbook, 'D2'), 5);
  assert.equal(valueAt(workbook, 'D3'), 15);

  history.undo();
  engine.rowsChanged(0);
  assert.deepEqual(names(workbook, engine.recalculate([])), ['S!D1', 'S!D2']);
  assert.equal(valueAt(workbook, 'D1'), 15);
  history.redo();
  engine.rowsChanged(0);
  engine.recalculate([]);
  assert.equal(valueAt(workbook, 'D1'), 5);
});

test('a change of hidden rows is a change of what the formula reads: the value of the file gives way', () => {
  // The file says 999 where the engine calculates 55: the value of the file stays ...
  const workbook = subtotals(999);
  const engine = new CalcEngine(workbook, NO_LIMIT);
  engine.recalculate();
  assert.equal(valueAt(workbook, 'B1'), 999);
  assert.deepEqual(cells(engine.mismatched), ['0!B1']);
  assert.equal(valueAt(workbook, 'B3'), 1998);
  // ... until rows are hidden.
  hide(workbook.sheets[0], [1]);
  engine.rowsChanged(0);
  engine.recalculate([]);
  assert.equal(valueAt(workbook, 'B1'), 54);
  assert.equal(valueAt(workbook, 'B3'), 108);
  assert.deepEqual(engine.mismatched, []);
});

test('a file whose subtotal was stored with rows filtered out agrees with the engine', () => {
  const workbook = subtotals(45);
  workbook.sheets[0].autoFilter = filterOn('A1:A10');
  hide(workbook.sheets[0], [10]);
  const engine = new CalcEngine(workbook, NO_LIMIT);
  engine.recalculate();
  assert.deepEqual(engine.mismatched, []);
  assert.equal(valueAt(workbook, 'B1'), 45);
});

test('rowsChanged() before the graph is built, and with a formula that reaches the function through a name', () => {
  const workbook = book({ S: { A1: 1, A2: 2, A3: 4, B1: { f: 'Shown', v: 7 }, B2: { f: 'SUBTOTAL(109,OFFSET(A1,0,0,3,1))', v: 7 }, B3: { f: 'SUM(A1:A3)', v: 7 } } }, [{ name: 'Shown', value: 'SUBTOTAL(109,S!$A$1:$A$3)' }]);
  const counted = counting('SUM');
  const engine = new CalcEngine(workbook, { ...NO_LIMIT, functions: counted.functions });
  hide(workbook.sheets[0], [2]);
  engine.rowsChanged(0);
  const changed = engine.recalculate([]);
  assert.deepEqual(names(workbook, changed), ['S!B1', 'S!B2']);
  assert.equal(valueAt(workbook, 'B1'), 5);
  assert.equal(valueAt(workbook, 'B2'), 5);
  assert.equal(valueAt(workbook, 'B3'), 7, 'not calculated');
  assert.equal(counted.calls.SUM, 0);
  // Asked for before invalidate(), answered by the recalculation after it.
  hide(workbook.sheets[0], [3]);
  engine.rowsChanged(0);
  engine.invalidate();
  engine.recalculate([]);
  assert.equal(valueAt(workbook, 'B1'), 1);
});

test('rowsChanged() while a recalculation is done in slices', () => {
  const data: Record<string, Spec> = {};
  const count = 600;
  for (let r = 1; r <= 50; r++) data[`A${r}`] = r;
  for (let r = 1; r <= count; r++) data[`B${r}`] = '=SUBTOTAL(109,A$1:A$50)';
  for (let r = 1; r <= count; r++) data[`C${r}`] = `=B${r}+1`;
  const workbook = book({ S: data });
  const engine = new CalcEngine(workbook, { budget: { work: 2000, milliseconds: 0 } });
  let calls = 0;
  engine.recalculate();
  while (engine.incomplete) {
    engine.recalculate([]);
    assert.ok(++calls < 10000);
  }
  assert.ok(calls > 3, `the recalculation took ${calls} more calls`);
  assert.equal(valueAt(workbook, `C${count}`), 1276);

  hide(workbook.sheets[0], [50]);
  engine.rowsChanged(0);
  engine.recalculate([]);
  assert.equal(engine.incomplete, true, 'one call is not enough');
  // Rows change again before the first change is calculated to its end.
  hide(workbook.sheets[0], [49]);
  engine.rowsChanged(0);
  calls = 0;
  while (engine.incomplete) {
    engine.recalculate([]);
    assert.ok(++calls < 10000);
  }
  for (const r of [1, 2, 300, count]) {
    assert.equal(valueAt(workbook, `B${r}`), 1176, `B${r}`);
    assert.equal(valueAt(workbook, `C${r}`), 1177, `C${r}`);
  }

  // The graph itself is built in slices: the request waits for it.
  let now = 0;
  const slow = new CalcEngine(workbook, { budget: { milliseconds: 5 }, clock: () => (now += 1) });
  hide(workbook.sheets[0], [48]);
  slow.rowsChanged(0);
  calls = 0;
  do {
    slow.recalculate([]);
    assert.ok(++calls < 100000);
  } while (slow.incomplete);
  assert.ok(calls > 3, `building and calculating took ${calls} calls`);
  assert.equal(valueAt(workbook, 'B1'), 1128);
  assert.equal(valueAt(workbook, `C${count}`), 1129);
});

test('previews read hidden rows too: evaluateText, evaluateShifted, batch', () => {
  const workbook = list();
  hide(workbook.sheets[0], [2, 3]);
  const engine = new CalcEngine(workbook, NO_LIMIT);
  assert.equal(engine.evaluateText('SUBTOTAL(109,A2:A6)', OUTSIDE), 28);
  // The formula of a conditional format written for B2, applied to B4: A2:A4 becomes A4:A6.
  assert.equal(engine.evaluateShifted('SUBTOTAL(109,A2:A4)', { sheet: 0, row: 1, col: 1 }, { row: 3, col: 1 }), 28);
  assert.equal(engine.evaluateShifted('SUBTOTAL(109,A2:A4)', { sheet: 0, row: 1, col: 1 }, { row: 1, col: 1 }), 4);
  assert.equal(engine.evaluateShifted('AGGREGATE(4,5,A$2:A2)', { sheet: 0, row: 1, col: 1 }, { row: 2, col: 1 }), 0);
  const both = engine.batch(() => [engine.evaluateText('SUBTOTAL(102,A2:A6)', OUTSIDE), engine.evaluateText('SUBTOTAL(2,A2:A6)', OUTSIDE)]);
  assert.deepEqual(both, [3, 5]);
  // The model changes between previews: nothing is remembered.
  hide(workbook.sheets[0], [2, 3], false);
  assert.equal(engine.evaluateText('SUBTOTAL(109,A2:A6)', OUTSIDE), 31);
});

test('a context without rowHidden counts every row', () => {
  const spec = FUNCTIONS.get('SUBTOTAL') as FunctionSpec;
  const values = [1, 2, 4];
  const context = {
    sheet: 0,
    row: 9,
    col: 9,
    date1904: false,
    cellValue: (_sheet: number, row: number) => values[row] ?? null,
    forEachCell: (_sheet: number, range: { top: number; bottom: number; left: number }, visit: (value: Scalar, row: number, col: number) => void) => {
      for (let r = range.top; r <= range.bottom && r < values.length; r++) visit(values[r], r, range.left);
    },
    usedExtent: () => ({ rows: 3, cols: 1 }),
    sheetIndex: () => undefined,
    sheetName: () => undefined,
    sheetCount: () => 1,
    now: () => 0,
    random: () => 0,
  };
  const ref = { ref: true as const, sheet: 0, range: { top: 0, left: 0, bottom: 2, right: 0 } };
  assert.equal(spec.call([109, ref], context), 7);
  assert.equal(spec.call([109, ref], { ...context, rowHidden: (_sheet: number, row: number) => row === 1 }), 5);
  assert.equal(spec.call([9, ref], { ...context, rowHidden: (_sheet: number, row: number) => row === 1 }), 7);
  assert.equal(spec.call([9, ref], { ...context, rowFiltered: (_sheet: number, row: number) => row === 1 }), 5);
});

test('rowsChanged() in a workbook of 100,000 formulas costs what the subtotals cost', () => {
  const data: Record<string, Spec> = {};
  const rows = 50000;
  for (let r = 1; r <= rows; r++) {
    data[`A${r}`] = r;
    data[`B${r}`] = `=A${r}*2`;
    data[`C${r}`] = `=B${r}+A${r}`;
  }
  for (let r = 1; r <= 20; r++) data[`E${r}`] = `=SUBTOTAL(109,A${r}:A${r + 99})`;
  const workbook = book({ S: data });
  const engine = new CalcEngine(workbook, NO_LIMIT);
  engine.recalculate();
  assert.equal(valueAt(workbook, 'E1'), 5050);
  hide(workbook.sheets[0], [5]);
  const started = Date.now();
  engine.rowsChanged(0);
  const changed = engine.recalculate([]);
  const took = Date.now() - started;
  assert.deepEqual(names(workbook, changed), ['S!E1', 'S!E2', 'S!E3', 'S!E4', 'S!E5']);
  assert.equal(valueAt(workbook, 'E1'), 5045);
  assert.ok(took < 500, `took ${took} ms`);
});

// ---------------------------------------------------------------------------
// evaluateRange

function ranges(): Workbook {
  return book(
    {
      S: { A1: 'a', A2: 'b', A3: 3, A5: true, B1: 10, B2: 20, B3: '#N/A', C1: { f: 'A3*2', v: 6 }, D1: 'L', E1: 2 },
      L: { A1: 'x', A2: 'y', A3: 'z', B2: 5 },
    },
    [
      { name: 'Letters', value: 'L!$A$1:$A$3' },
      { name: 'Local', value: 'S!$B$1:$B$2', localSheet: 0 },
      { name: 'Local', value: 'L!$B$2', localSheet: 1 },
      { name: 'Twice', value: 'S!$B$1*2' },
    ],
  );
}

const range = (value: RangeValues | CalcError): RangeValues => {
  assert.ok(!isError(value), `an error: ${isError(value) ? value.error : ''}`);
  return value as RangeValues;
};

test('evaluateRange gives every value of a range, with empty cells and errors in place', () => {
  const engine = new CalcEngine(ranges(), NO_LIMIT);
  const found = range(engine.evaluateRange('A1:B3', OUTSIDE));
  assert.equal(found.rows, 3);
  assert.equal(found.cols, 2);
  assert.equal(found.truncated, false);
  assert.deepEqual(found.values.map(code), ['a', 10, 'b', 20, 3, '#N/A']);
  assert.equal(found.sheet, 0);
  assert.deepEqual(found.range, { top: 0, left: 0, bottom: 2, right: 1 });
  // An empty cell is null, not 0.
  assert.deepEqual(range(engine.evaluateRange('$A$1:$A$5', OUTSIDE)).values, ['a', 'b', 3, null, true]);
  // A formula cell answers with the value the model holds.
  assert.deepEqual(range(engine.evaluateRange('C1', OUTSIDE)).values, [6]);
  assert.deepEqual(range(engine.evaluateRange('=A1:A2', OUTSIDE)).values, ['a', 'b'], 'a leading = is allowed as in evaluateText');
});

test('evaluateRange reads other sheets, defined names, OFFSET and INDIRECT', () => {
  const engine = new CalcEngine(ranges(), NO_LIMIT);
  const other = range(engine.evaluateRange('L!$A$1:$A$500', OUTSIDE));
  assert.deepEqual(other.values, ['x', 'y', 'z']);
  assert.equal(other.rows, 3, 'cut to the used part of the sheet');
  assert.equal(other.sheet, 1);
  assert.deepEqual(range(engine.evaluateRange('Letters', OUTSIDE)).values, ['x', 'y', 'z']);
  assert.deepEqual(range(engine.evaluateRange('Local', OUTSIDE)).values, [10, 20]);
  assert.deepEqual(range(engine.evaluateRange('Local', { sheet: 1, row: 9, col: 9 })).values, [5], 'the name of the sheet the formula is in');
  assert.deepEqual(range(engine.evaluateRange('OFFSET($A$1,0,0,COUNTA($A:$A),1)', OUTSIDE)).values, ['a', 'b', 3, null]);
  assert.deepEqual(range(engine.evaluateRange('OFFSET(L!A1,1,0,2,2)', OUTSIDE)).values, ['y', 5, 'z', null]);
  assert.deepEqual(range(engine.evaluateRange('INDIRECT(D1&"!A"&E1&":A3")', OUTSIDE)).values, ['y', 'z']);
  assert.deepEqual(range(engine.evaluateRange('INDIRECT("Letters")', OUTSIDE)).values, ['x', 'y', 'z']);
});

test('evaluateRange gives a single cell, a scalar and an array as they are, and errors as errors', () => {
  const engine = new CalcEngine(ranges(), NO_LIMIT);
  const single = range(engine.evaluateRange('B2', OUTSIDE));
  assert.deepEqual([single.rows, single.cols, single.values], [1, 1, [20]]);
  assert.deepEqual(range(engine.evaluateRange('L!C9', OUTSIDE)).values, [null], 'an empty cell outside the used part');
  const scalar = range(engine.evaluateRange('1+2', OUTSIDE));
  assert.deepEqual([scalar.rows, scalar.cols, scalar.values, scalar.sheet], [1, 1, [3], undefined]);
  assert.deepEqual(range(engine.evaluateRange('Twice', OUTSIDE)).values, [20]);
  const array = range(engine.evaluateRange('{1,2;3,"x"}', OUTSIDE));
  assert.deepEqual([array.rows, array.cols, array.values], [2, 2, [1, 2, 3, 'x']]);
  // A cell that holds an error is a value of the range; a formula that is an error is the error.
  assert.deepEqual(range(engine.evaluateRange('B3', OUTSIDE)).values.map(code), ['#N/A']);
  assert.equal(code(engine.evaluateRange('1/0', OUTSIDE)), '#DIV/0!');
  assert.equal(code(engine.evaluateRange('Nothing', OUTSIDE)), '#NAME?');
  assert.equal(code(engine.evaluateRange('INDIRECT("no such")', OUTSIDE)), '#REF!');
  assert.equal(code(engine.evaluateRange('OFFSET(A1,-1,0)', OUTSIDE)), '#REF!');
  // Never throws.
  assert.equal(code(engine.evaluateRange('A1:', OUTSIDE)), '#VALUE!');
  assert.equal(code(engine.evaluateRange('', OUTSIDE)), '#VALUE!');
  assert.equal(code(engine.evaluateRange(undefined as unknown as string, undefined as unknown as CellAddress)), '#VALUE!');
  assert.deepEqual(range(engine.evaluateRange('A1', { sheet: 99, row: -1, col: NaN })).values, ['a']);
});

test('evaluateRange moves relative references by the distance from the anchor', () => {
  const engine = new CalcEngine(ranges(), NO_LIMIT);
  // Written for C1, calculated for C2: A1:B1 becomes A2:B2; absolute parts stay.
  assert.deepEqual(range(engine.evaluateRange('A1:B1', { sheet: 0, row: 1, col: 2 }, { anchor: { row: 0, col: 2 } })).values, ['b', 20]);
  assert.deepEqual(range(engine.evaluateRange('$A$1:B1', { sheet: 0, row: 1, col: 2 }, { anchor: { row: 0, col: 2 } })).values.map(code), ['a', 10, 'b', 20]);
  assert.deepEqual(range(engine.evaluateRange('A1:B1', { sheet: 0, row: 1, col: 2 })).values, ['a', 10], 'no anchor: as written');
  assert.equal(code(engine.evaluateRange('A1:B1', A1, { anchor: { row: 1, col: 0 } })), '#REF!', 'pushed off the sheet');
});

test('evaluateRange cuts a result that is too large and says so', () => {
  const data: Record<string, Spec> = {};
  for (let r = 1; r <= 3000; r++) {
    data[`A${r}`] = r;
    data[`B${r}`] = -r;
  }
  const workbook = book({ S: data });
  const engine = new CalcEngine(workbook, NO_LIMIT);
  const whole = range(engine.evaluateRange('A:B', { sheet: 0, row: 0, col: 5 }));
  assert.equal(whole.rows, 3000);
  assert.equal(whole.values.length, 6000);
  assert.equal(whole.truncated, false);
  const cut = range(engine.evaluateRange('A:B', { sheet: 0, row: 0, col: 5 }, { maxCells: 5 }));
  assert.deepEqual([cut.rows, cut.cols, cut.truncated], [3000, 2, true]);
  assert.deepEqual(cut.values, [1, -1, 2, -2, 3]);
  assert.deepEqual(range(engine.evaluateRange('A:B', { sheet: 0, row: 0, col: 5 }, { maxCells: 6000 })).truncated, false);
  // A limit that is no limit at all gives one value; the whole sheet is cut to what is in use.
  assert.deepEqual(range(engine.evaluateRange('A:B', { sheet: 0, row: 0, col: 5 }, { maxCells: 0 })).values, [1]);
  assert.deepEqual(range(engine.evaluateRange('A:B', { sheet: 0, row: 0, col: 5 }, { maxCells: -3 })).values, [1]);
  assert.equal(range(engine.evaluateRange('1:1048576', { sheet: 0, row: 0, col: 5 }, { maxCells: 1e12 })).values.length, 6000);
  const array = range(engine.evaluateRange('SEQUENCE(100)', OUTSIDE, { maxCells: 3 }));
  assert.deepEqual([array.rows, array.values, array.truncated], [100, [1, 2, 3], true]);
});

test('evaluateRange keeps to the budget and shares it within batch()', () => {
  const data: Record<string, Spec> = {};
  for (let r = 1; r <= 5000; r++) data[`A${r}`] = r;
  const workbook = book({ S: data });
  const engine = new CalcEngine(workbook, { budget: { work: 12000, milliseconds: 0 } });
  // 5,000 cells fit, each call with a budget of its own ...
  assert.equal(range(engine.evaluateRange('A1:A5000', { sheet: 0, row: 0, col: 3 })).values.length, 5000);
  assert.equal(range(engine.evaluateRange('A1:A5000', { sheet: 0, row: 0, col: 3 })).values.length, 5000);
  assert.equal(range(engine.evaluateRange('A1:A5000', { sheet: 0, row: 0, col: 3 })).values.length, 5000);
  // ... and three of them do not fit into one.
  const results = engine.batch(() => [1, 2, 3].map(() => engine.evaluateRange('A1:A5000', { sheet: 0, row: 0, col: 3 })));
  assert.deepEqual(results.map((result) => (isError(result) ? result.error : result.values.length)), [5000, 5000, '#VALUE!']);
  // A preview does not disturb the model or what the engine lists.
  assert.equal(engine.incomplete, false);
  assert.equal(valueAt(workbook, 'A1'), 1);
});

test('the list of a data validation reads a formula source in one call', () => {
  const data: Record<string, Spec> = { C1: 'S', C2: 3 };
  const rows = 25000;
  for (let r = 1; r <= rows; r++) data[`A${r}`] = r % 7 === 0 ? null : `item ${r}`;
  const workbook = book({ S: data });
  const rule = (formula1: string, ref: string): DataValidation => {
    const at = parseCellRef(ref) as { row: number; col: number };
    return { ranges: [{ top: at.row, left: at.col, bottom: at.row + 1, right: at.col }], type: 'list', operator: 'between', formula1, allowBlank: true, hideDropDown: false, showInputMessage: false, showErrorMessage: true, errorStyle: 'stop' };
  };
  workbook.sheets[0].dataValidations = [rule('INDIRECT($C$1&"!A:A")', 'E1'), rule('OFFSET($A$1,0,0,$C$2,1)', 'F1'), rule('OFFSET(A1,0,0,2,1)', 'G1'), rule('INDIRECT("no such")', 'H1')];
  // The source used to be asked for one value at a time, with INDEX.
  const counted = counting('INDEX');
  const engine = new CalcEngine(workbook, { ...NO_LIMIT, functions: counted.functions });
  const rules = new DataValidations(workbook, workbook.sheets[0], { engine: () => engine });

  const long = rules.list(0, 4);
  assert.ok(long);
  assert.equal(long.items.length, MAX_LIST_ITEMS);
  assert.equal(long.truncated, true);
  assert.deepEqual(long.items.slice(0, 7), ['item 1', 'item 2', 'item 3', 'item 4', 'item 5', 'item 6', 'item 8'], 'empty cells are passed over');
  assert.equal(counted.calls.INDEX, 0);

  assert.deepEqual(rules.list(0, 5), { items: ['item 1', 'item 2', 'item 3'], truncated: false });
  // Relative references move with the cell: the rule is written for G1, and G2 reads A2:A3.
  assert.deepEqual(rules.list(0, 6), { items: ['item 1', 'item 2'], truncated: false });
  assert.deepEqual(rules.list(1, 6), { items: ['item 2', 'item 3'], truncated: false });
  assert.equal(rules.list(0, 7), undefined, 'a source that is an error cannot be read');
  assert.equal(counted.calls.INDEX, 0);
});

// ---------------------------------------------------------------------------
// cycles

test('cycles() lists the circular references the formulas name, before anything is calculated', () => {
  const workbook = book({
    S: {
      A1: { f: 'B1+1', v: 1 },
      B1: { f: 'A1+1', v: 2 },
      C1: { f: 'C1+1', v: 3 },
      D1: { f: 'IF(TRUE,1,E1)', v: 1 },
      E1: { f: 'SUM(D1:D3)', v: 1 },
      F1: { f: 'A1+C1', v: 4 },
      G1: { f: 'ROW(G1)', v: 1 },
    },
    T: { A1: { f: 'S!A1+T!B1', v: 0 }, B1: { f: 'A1', v: 0 } },
  });
  const engine = new CalcEngine(workbook, NO_LIMIT);
  assert.equal(engine.graphReady, false);
  assert.deepEqual(engine.cycles(), [], 'not known yet');
  // Nothing changed, so nothing is calculated: `circular` stays empty.
  assert.deepEqual(engine.recalculate([]), []);
  assert.deepEqual(engine.circular, []);
  assert.equal(engine.graphReady, true);
  assert.deepEqual(cells(engine.cycles()), ['0!A1', '0!B1', '0!C1', '0!D1', '0!E1', '1!A1', '1!B1']);
  assert.equal(valueAt(workbook, 'A1'), 1, 'the values of the file stay');

  // An edit that breaks a circle, one that closes another.
  const changed = [edit(workbook, 0, 'B1', 5), edit(workbook, 0, 'F1', '=G1'), edit(workbook, 0, 'G1', '=F1')];
  engine.recalculate(changed);
  assert.deepEqual(cells(engine.cycles()), ['0!C1', '0!D1', '0!E1', '0!F1', '0!G1', '1!A1', '1!B1']);
  assert.deepEqual(cells(engine.circular), ['0!F1', '0!G1', '1!A1', '1!B1'], 'only what was calculated: C1, D1 and E1 were not');
  engine.recalculate([edit(workbook, 0, 'C1', undefined), edit(workbook, 1, 'B1', 1)]);
  assert.deepEqual(cells(engine.cycles()), ['0!D1', '0!E1', '0!F1', '0!G1']);

  // The answer is a copy.
  engine.cycles().pop();
  assert.equal(engine.cycles().length, 4);
  engine.invalidate();
  assert.deepEqual(engine.cycles(), []);
  assert.equal(engine.graphReady, false);
});

test('cycles() is empty until a graph built in slices is complete', () => {
  const data: Record<string, Spec> = {};
  const count = 3000;
  for (let r = 1; r <= count; r++) data[`A${r}`] = { f: `A${r === count ? 1 : r + 1}+1`, v: 0 };
  data.C1 = { f: 'C2', v: 0 };
  data.C2 = { f: 'C1', v: 0 };
  const workbook = book({ S: data });
  let now = 0;
  const engine = new CalcEngine(workbook, { budget: { milliseconds: 2 }, clock: () => (now += 1) });
  let calls = 0;
  let early = 0;
  do {
    engine.recalculate([]);
    if (!engine.graphReady) {
      early++;
      assert.deepEqual(engine.cycles(), [], 'part of the answer is not given');
      assert.equal(engine.incomplete, true);
    }
    assert.ok(++calls < 100000);
  } while (engine.incomplete);
  assert.ok(early > 3, `the graph took ${early} calls`);
  assert.equal(engine.graphReady, true);
  assert.equal(engine.cycles().length, count + 2);
});

// ---------------------------------------------------------------------------
// namesChanged

function named(): Workbook {
  return book(
    {
      S: {
        A1: 1,
        A2: 2,
        A3: 4,
        B1: { f: 'SUM(Part)', v: 7 },
        B2: { f: 'SUM(A1:A3)', v: 7 },
        B3: { f: 'SUM(Other)', v: 4 },
        B4: { f: 'Total*10', v: 70 },
        B5: { f: 'SUM(Later)', v: 123 },
        B6: { f: 'B1+1', v: 8 },
      },
    },
    [
      { name: 'Part', value: 'S!$A$1:$A$3' },
      { name: 'Other', value: 'S!$A$3' },
      { name: 'Total', value: 'SUM(Part)' },
    ],
  );
}

const rename = (workbook: Workbook, name: string, value: string | undefined): void => {
  const others = workbook.definedNames.filter((entry) => entry.name !== name);
  workbook.definedNames = value === undefined ? others : others.concat([{ name, value }]);
};

test('namesChanged() recalculates the formulas that use a name that changed, and only those', () => {
  const workbook = named();
  const counted = counting('SUM');
  const engine = new CalcEngine(workbook, { ...NO_LIMIT, functions: counted.functions });
  engine.recalculate();
  assert.deepEqual(cells(engine.unsupported), ['0!B5'], 'a name the workbook does not have');
  assert.equal(valueAt(workbook, 'B5'), 123);

  counted.reset();
  rename(workbook, 'Part', 'S!$A$1:$A$2');
  engine.namesChanged();
  assert.equal(counted.calls.SUM, 0);
  const changed = engine.recalculate([]);
  assert.deepEqual(names(workbook, changed), ['S!B1', 'S!B4', 'S!B6']);
  assert.equal(valueAt(workbook, 'B1'), 3);
  assert.equal(valueAt(workbook, 'B4'), 30, 'through another name');
  assert.equal(valueAt(workbook, 'B6'), 4, 'what reads such a formula follows');
  assert.equal(counted.calls.SUM, 2, 'SUM(A1:A3) and SUM(Other) were left alone');

  // The graph follows the name: A3 is no longer read, A1 still is.
  assert.deepEqual(engine.recalculate([edit(workbook, 0, 'A3', 100)]).length, 1);
  assert.equal(valueAt(workbook, 'B1'), 3);
  assert.equal(valueAt(workbook, 'B2'), 103);
  engine.recalculate([edit(workbook, 0, 'A1', 10)]);
  assert.equal(valueAt(workbook, 'B1'), 12);
  assert.equal(valueAt(workbook, 'B4'), 120);

  // And previews see it at once.
  rename(workbook, 'Part', 'S!$A$2');
  engine.namesChanged();
  assert.equal(engine.evaluateText('SUM(Part)', OUTSIDE), 2);
  engine.recalculate([]);
  assert.equal(valueAt(workbook, 'B1'), 2);
});

test('namesChanged() with nothing changed, or with a name no formula uses, calculates nothing', () => {
  const workbook = named();
  const counted = counting('SUM');
  const engine = new CalcEngine(workbook, { ...NO_LIMIT, functions: counted.functions });
  engine.recalculate();
  counted.reset();
  engine.namesChanged();
  assert.deepEqual(engine.recalculate([]), []);
  // The filter of a sheet comes and goes with a hidden name of that sheet.
  workbook.definedNames = workbook.definedNames.concat([{ name: '_xlnm._FilterDatabase', value: 'S!$A$1:$A$3', localSheet: 0, hidden: true }]);
  engine.namesChanged();
  assert.deepEqual(engine.recalculate([]), []);
  workbook.definedNames = workbook.definedNames.filter((entry) => entry.name !== '_xlnm._FilterDatabase');
  engine.namesChanged();
  assert.deepEqual(engine.recalculate([]), []);
  // The same names in a new array, in another order.
  workbook.definedNames = workbook.definedNames.slice().reverse();
  engine.namesChanged();
  assert.deepEqual(engine.recalculate([]), []);
  assert.equal(counted.calls.SUM, 0);
  assert.equal(valueAt(workbook, 'B1'), 7);
});

test('namesChanged() after a name was added or removed', () => {
  const workbook = named();
  const engine = new CalcEngine(workbook, NO_LIMIT);
  engine.recalculate();
  // A name that was missing: the formula is calculated from now on.
  rename(workbook, 'Later', 'S!$A$2:$A$3');
  engine.namesChanged();
  assert.deepEqual(names(workbook, engine.recalculate([])), ['S!B5']);
  assert.equal(valueAt(workbook, 'B5'), 6);
  assert.deepEqual(engine.unsupported, []);
  engine.recalculate([edit(workbook, 0, 'A2', 20)]);
  assert.equal(valueAt(workbook, 'B5'), 24);

  // A name that is removed: the value the engine wrote is not one of the file, so the error shows.
  rename(workbook, 'Later', undefined);
  engine.namesChanged();
  engine.recalculate([]);
  assert.equal(valueAt(workbook, 'B5'), '#NAME?');
  // A value of the file stays when the name it needs goes.
  rename(workbook, 'Other', undefined);
  engine.namesChanged();
  engine.recalculate([]);
  assert.equal(valueAt(workbook, 'B3'), 4);
  assert.deepEqual(cells(engine.unsupported), ['0!B3']);

  // A name of one sheet in front of the name of the workbook.
  workbook.definedNames = workbook.definedNames.concat([{ name: 'part', value: 'S!$A$3', localSheet: 0 }]);
  engine.namesChanged();
  engine.recalculate([]);
  assert.equal(valueAt(workbook, 'B1'), 4);
  workbook.definedNames = workbook.definedNames.filter((entry) => entry.localSheet === undefined);
  engine.namesChanged();
  engine.recalculate([]);
  assert.equal(valueAt(workbook, 'B1'), 25);
});

test('namesChanged() replaces a value of the file that the old name gave', () => {
  const workbook = named();
  // The file says 7 for B1; with the new name the engine says 6 and that is what counts.
  const engine = new CalcEngine(workbook, NO_LIMIT);
  engine.recalculate([]);
  rename(workbook, 'Part', 'S!$A$2:$A$3');
  engine.namesChanged();
  engine.recalculate([]);
  assert.equal(valueAt(workbook, 'B1'), 6);
  assert.deepEqual(engine.mismatched, []);
  assert.equal(valueAt(workbook, 'B2'), 7);
});

test('namesChanged() before and while the graph is built, and circles through names', () => {
  const workbook = named();
  const fresh = new CalcEngine(workbook, NO_LIMIT);
  fresh.namesChanged();
  rename(workbook, 'Part', 'S!$A$1');
  // Only prepared (a preview), not built.
  assert.equal(fresh.evaluateText('SUM(Part)', OUTSIDE), 1);
  rename(workbook, 'Part', 'S!$A$2');
  fresh.namesChanged();
  assert.equal(fresh.evaluateText('SUM(Part)', OUTSIDE), 2);

  // While the graph is built in slices: the formulas read so far are read again.
  const data: Record<string, Spec> = { A1: 1, A2: 2 };
  for (let r = 1; r <= 2000; r++) data[`B${r}`] = '=SUM(Part)';
  const large = book({ S: data }, [{ name: 'Part', value: 'S!$A$1' }]);
  let now = 0;
  const engine = new CalcEngine(large, { budget: { milliseconds: 2 }, clock: () => (now += 1) });
  engine.recalculate();
  assert.equal(engine.incomplete, true);
  assert.equal(engine.graphReady, false);
  rename(large, 'Part', 'S!$A$1:$A$2');
  engine.namesChanged();
  assert.equal(engine.incomplete, true);
  let calls = 0;
  while (engine.incomplete) {
    engine.recalculate([]);
    assert.ok(++calls < 100000);
  }
  assert.equal(valueAt(large, 'B1'), 3);
  assert.equal(valueAt(large, 'B2000'), 3);
  engine.recalculate([edit(large, 0, 'A2', 5)]);
  while (engine.incomplete) engine.recalculate([]);
  assert.equal(valueAt(large, 'B2000'), 6);

  // A name that now points at the formula that uses it.
  const loop = book({ S: { A1: 1, B1: { f: 'Part+1', v: 2 } } }, [{ name: 'Part', value: 'S!$A$1' }]);
  const third = new CalcEngine(loop, NO_LIMIT);
  third.recalculate();
  assert.deepEqual(third.cycles(), []);
  rename(loop, 'Part', 'S!$B$1');
  third.namesChanged();
  third.recalculate([]);
  assert.deepEqual(cells(third.cycles()), ['0!B1']);
  assert.equal(valueAt(loop, 'B1'), 0);
});
