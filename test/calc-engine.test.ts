import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine, type CellAddress } from '../src/calc/engine';
import { FormulaParseError } from '../src/calc/parse';
import { isArray, isError, isRef, type CalcContext, type FunctionSpec, type Scalar, type Value } from '../src/calc/values';
import { open } from '../src/index';
import {
  cellName,
  createSheet,
  emptyStyleSheet,
  getCell,
  parseCellRef,
  type Cell,
  type CellFormula,
  type DefinedName,
  type Row,
  type Sheet,
  type Workbook,
} from '../src/model/sheet/types';
import { applyOp } from '../src/ops/ops';
import { buildXlsx, worksheetXml } from './fixtures';

// ---------------------------------------------------------------------------
// Building workbooks by hand

type Plain = number | string | boolean | null;
interface FormulaSpec {
  f: string;
  /** Value the file would have stored for the formula. */
  v?: Plain;
  type?: CellFormula['type'];
  ref?: string;
  si?: number;
  error?: boolean;
}
type Spec = Plain | FormulaSpec;

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
    if (spec.si !== undefined) formula.si = spec.si;
    return { kind: kindOf(value, spec.error), value, style: 0, formula };
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

function book(sheets: Record<string, Record<string, Spec>>, definedNames: DefinedName[] = []): Workbook {
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

function at(workbook: Workbook, address: string, sheet = 0): Cell | undefined {
  const cell = parseCellRef(address) as { row: number; col: number };
  return getCell(workbook.sheets[sheet], cell.row, cell.col);
}

const valueAt = (workbook: Workbook, address: string, sheet = 0) => at(workbook, address, sheet)?.value;

/** The cells a recalculation reports, as "Sheet!A1" names in order. */
function names(workbook: Workbook, result: { sheet: number; cells: { row: number; col: number }[] }[]): string[] {
  const out: string[] = [];
  for (const entry of result) for (const cell of entry.cells) out.push(`${workbook.sheets[entry.sheet].name}!${cellName(cell.row, cell.col)}`);
  return out.sort();
}

// ---------------------------------------------------------------------------
// A small function library, so that these tests do not depend on the real one

function flatten(args: Value[], context: CalcContext): Scalar[] {
  const out: Scalar[] = [];
  for (const arg of args) {
    if (isRef(arg)) context.forEachCell(arg.sheet, arg.range, (value) => out.push(value));
    else if (isArray(arg)) out.push(...arg.data);
    else out.push(arg);
  }
  return out;
}

const counters = { tick: 0 };

const LIBRARY = new Map<string, FunctionSpec>([
  [
    'SUM',
    {
      min: 1,
      call: (args, context) => {
        let total = 0;
        for (const value of flatten(args, context)) {
          if (isError(value)) return value;
          if (typeof value === 'number') total += value;
        }
        return total;
      },
    },
  ],
  ['MAX', { min: 1, call: (args, context) => flatten(args, context).reduce<number>((best, value) => (typeof value === 'number' && value > best ? value : best), -Infinity) }],
  ['CONCAT', { min: 1, call: (args, context) => flatten(args, context).map((value) => (value === null ? '' : String(value))).join('') }],
  ['NOW', { min: 0, max: 0, call: (_args, context) => context.now() }],
  ['RAND', { min: 0, max: 0, call: (_args, context) => context.random() }],
  ['TICK', { min: 0, call: () => ++counters.tick }],
  ['BOOM', { min: 0, call: () => { throw new Error('broken function'); } }],
  ['OVERFLOW', { min: 0, call: () => { const dive = (): number => dive() + 1; return dive(); } }],
  ['DATE1904', { min: 0, call: (_args, context) => context.date1904 }],
]);

const engineFor = (workbook: Workbook, options = {}) => new CalcEngine(workbook, { functions: LIBRARY, ...options });

// ---------------------------------------------------------------------------

test('a full recalculation works out every formula and writes the results', () => {
  const workbook = book({
    S: {
      A1: 2,
      A2: 3,
      B1: '=A1+A2',
      B2: '=B1*10',
      B3: '=A1&"x"',
      B4: '=A1>A2',
      B5: '=1/0',
      B6: '=Z99',
      B7: '=SUM(A1:A2,B1)',
      B8: { f: 'A1*2', v: 4 },
    },
  });
  const untouched = at(workbook, 'B8');
  const engine = engineFor(workbook);
  const result = engine.recalculate();
  assert.deepEqual(names(workbook, result), ['S!B1', 'S!B2', 'S!B3', 'S!B4', 'S!B5', 'S!B6', 'S!B7']);
  assert.deepEqual([at(workbook, 'B1')?.kind, valueAt(workbook, 'B1')], ['number', 5]);
  assert.equal(valueAt(workbook, 'B2'), 50);
  assert.deepEqual([at(workbook, 'B3')?.kind, valueAt(workbook, 'B3')], ['string', '2x']);
  assert.deepEqual([at(workbook, 'B4')?.kind, valueAt(workbook, 'B4')], ['boolean', false]);
  assert.deepEqual([at(workbook, 'B5')?.kind, valueAt(workbook, 'B5')], ['error', '#DIV/0!']);
  // A reference to an empty cell shows 0.
  assert.deepEqual([at(workbook, 'B6')?.kind, valueAt(workbook, 'B6')], ['number', 0]);
  assert.equal(valueAt(workbook, 'B7'), 10);
  // The formulas stay.
  assert.equal(at(workbook, 'B1')?.formula?.text, 'A1+A2');
  // A value that was already right is not written again.
  assert.equal(at(workbook, 'B8'), untouched);
  assert.deepEqual(engine.recalculate(), []);
  assert.deepEqual(engine.circular, []);
  assert.deepEqual(engine.unparsed, []);
});

test('only what depends on a change is recalculated; stored values are trusted', () => {
  const workbook = book({
    S: {
      A1: 1,
      A2: 10,
      B1: { f: 'A1+1', v: 2 },
      C1: { f: 'B1*2', v: 4 },
      D1: { f: 'SUM(A1:C1)', v: 7 },
      // Depends on A2 only. The stored value is not what the formula gives: it must survive.
      B2: { f: 'A2+1', v: 999 },
      C2: { f: 'B2+1', v: 1000 },
      E1: { f: 'D1+C2', v: 1007 },
    },
  });
  const engine = engineFor(workbook);
  const result = engine.recalculate([edit(workbook, 0, 'A1', 5)]);
  assert.deepEqual(names(workbook, result), ['S!B1', 'S!C1', 'S!D1', 'S!E1']);
  assert.equal(valueAt(workbook, 'B1'), 6);
  assert.equal(valueAt(workbook, 'C1'), 12);
  assert.equal(valueAt(workbook, 'D1'), 23);
  assert.equal(valueAt(workbook, 'B2'), 999);
  assert.equal(valueAt(workbook, 'C2'), 1000);
  // E1 reads C2 as the model has it.
  assert.equal(valueAt(workbook, 'E1'), 1023);

  // A change that changes no result reports nothing.
  assert.deepEqual(engine.recalculate([edit(workbook, 0, 'A1', 5)]), []);
  // Changing A2 now brings its own dependants up to date.
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'A2', 20)])), ['S!B2', 'S!C2', 'S!E1']);
  assert.equal(valueAt(workbook, 'E1'), 45);
  // Clearing a cell counts as a change.
  engine.recalculate([edit(workbook, 0, 'A1', undefined)]);
  assert.equal(valueAt(workbook, 'B1'), 1);
});

test('cells and rows are replaced, never altered', () => {
  const workbook = book({ S: { A1: 1, B1: { f: 'A1+1', v: 2 }, C1: 'keep', A2: 5, B2: { f: 'A2', v: 5 } } });
  const sheet = workbook.sheets[0];
  const oldRow = sheet.rows.get(0) as Row;
  const oldCell = oldRow.cells[1];
  const otherRow = sheet.rows.get(1);
  const text = oldRow.cells[2];
  const engine = engineFor(workbook);
  const edited = edit(workbook, 0, 'A1', 41);
  const rowAfterEdit = sheet.rows.get(0) as Row;
  engine.recalculate([edited]);

  // The old objects are as they were.
  assert.deepEqual(oldCell, { kind: 'number', value: 2, style: 0, formula: { text: 'A1+1', type: 'normal' } });
  assert.equal(rowAfterEdit.cells[1], oldCell);
  const newRow = sheet.rows.get(0) as Row;
  assert.notEqual(newRow, rowAfterEdit);
  assert.notEqual(newRow.cells[1], oldCell);
  assert.equal(newRow.cells[1].value, 42);
  assert.equal(newRow.cells[1].formula, oldCell.formula);
  // Cells that did not change are shared with the old row; rows without a change are the same objects.
  assert.equal(newRow.cells[2], text);
  assert.equal(sheet.rows.get(1), otherRow);
  assert.deepEqual(Array.from(sheet.changedRows || []), [0]);
  assert.equal(sheet.modified, true);

  // Each recalculation copies a row anew: the row of the previous one may be held for undo.
  engine.recalculate([edit(workbook, 0, 'A1', 1)]);
  assert.equal(newRow.cells[1].value, 42);
  assert.equal(valueAt(workbook, 'B1'), 2);
});

test('a result of another kind replaces what the cell held', () => {
  const workbook = book({ S: { A1: 1, B1: { f: 'IF(A1=1,"one",A1)', v: 'x' } } });
  const sheet = workbook.sheets[0];
  const row = sheet.rows.get(0) as Row;
  // As read from a file: text from the shared strings, with formatting runs, in a styled cell.
  row.cells[1] = { ...row.cells[1], runs: [{ text: 'x' }], si: 3, style: 7, cm: 1 };
  const engine = engineFor(workbook);
  // The value of the file ("x") is not what the formula gives: it is only replaced when that is asked for.
  engine.recalculate(undefined, { overwrite: true });
  assert.deepEqual(at(workbook, 'B1'), { kind: 'string', value: 'one', style: 7, cm: 1, formula: { text: 'IF(A1=1,"one",A1)', type: 'normal' } });
  engine.recalculate([edit(workbook, 0, 'A1', 2)]);
  assert.deepEqual([at(workbook, 'B1')?.kind, valueAt(workbook, 'B1'), at(workbook, 'B1')?.style], ['number', 2, 7]);
});

test('what plain cells are worth', () => {
  const workbook = book({
    S: {
      B1: '=A2+1',
      B2: '=A3',
      B3: '=A4+1',
      B4: '=A5&"!"',
      B5: '=A6',
      B6: '=A7=""',
      B7: '=A8',
    },
  });
  const sheet = workbook.sheets[0];
  place(sheet, 1, 0, { kind: 'date', value: '2024-01-01T12:00:00', style: 0 });
  place(sheet, 2, 0, { kind: 'error', value: '#N/A', style: 0 });
  place(sheet, 3, 0, { kind: 'boolean', value: true, style: 0 });
  place(sheet, 4, 0, { kind: 'date', value: 'not a date', style: 0 });
  place(sheet, 5, 0, { kind: 'error', value: '#WHATEVER', style: 0 });
  place(sheet, 6, 0, { kind: 'blank', value: null, style: 3 });
  place(sheet, 7, 0, { kind: 'number', value: 'oops' as unknown as number, style: 0 });
  engineFor(workbook).recalculate();
  assert.equal(valueAt(workbook, 'B1'), 45292.5 + 1);
  assert.deepEqual([at(workbook, 'B2')?.kind, valueAt(workbook, 'B2')], ['error', '#N/A']);
  assert.equal(valueAt(workbook, 'B3'), 2);
  assert.equal(valueAt(workbook, 'B4'), 'not a date!');
  assert.equal(valueAt(workbook, 'B5'), '#VALUE!');
  assert.equal(valueAt(workbook, 'B6'), true);
  assert.equal(valueAt(workbook, 'B7'), 0);
});

test('ranges, whole columns and other sheets', () => {
  const workbook = book({
    Data: { A1: 1, A2: 2, A3: 3, B1: 10, C5: '=SUM(A1:A3)', C6: '=SUM(A:A)', C7: '=SUM(1:1)', C8: '=SUM(A1:B3)' },
    'Sum Sheet': { A1: '=Data!A1+1', A2: "=SUM(Data!A1:A3)*2", A3: '=A1+A2', A4: "='Sum Sheet'!A3", A5: '=SUM(data!$A:$A)' },
  });
  const engine = engineFor(workbook);
  engine.recalculate();
  assert.deepEqual(['C5', 'C6', 'C7', 'C8'].map((cell) => valueAt(workbook, cell)), [6, 6, 11, 16]);
  assert.deepEqual(['A1', 'A2', 'A3', 'A4', 'A5'].map((cell) => valueAt(workbook, cell, 1)), [2, 12, 14, 14, 6]);

  const result = engine.recalculate([edit(workbook, 0, 'A2', 20)]);
  assert.deepEqual(names(workbook, result), ['Data!C5', 'Data!C6', 'Data!C8', 'Sum Sheet!A2', 'Sum Sheet!A3', 'Sum Sheet!A4', 'Sum Sheet!A5']);
  assert.deepEqual(['C5', 'C6', 'C7', 'C8'].map((cell) => valueAt(workbook, cell)), [24, 24, 11, 34]);
  assert.deepEqual(['A1', 'A2', 'A3', 'A4', 'A5'].map((cell) => valueAt(workbook, cell, 1)), [2, 48, 50, 50, 24]);

  // A new cell far down a whole-column reference.
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'A5000', 100)])), ['Data!C6', 'Sum Sheet!A5']);
  assert.equal(valueAt(workbook, 'C6'), 124);
  // And one in the row a whole-row reference covers.
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'ZZ1', 5)])), ['Data!C7']);
  assert.equal(valueAt(workbook, 'C7'), 16);
});

test('ranges of every shape find their dependants', () => {
  // Tall, wide, both, and small ranges, each read by one formula on another sheet.
  const workbook = book({
    Data: { A1: 0 },
    F: {
      A1: '=SUM(Data!B2:B5000)',
      A2: '=SUM(Data!B10:ZZ12)',
      A3: '=SUM(Data!D100:ZZ9000)',
      A4: '=SUM(Data!D4:E6)',
      A5: '=SUM(Data!B2:C3000)+SUM(Data!B2:C3000)',
    },
  });
  const engine = engineFor(workbook);
  engine.recalculate();
  const change = (address: string) => names(workbook, engine.recalculate([edit(workbook, 0, address, 1)]));
  assert.deepEqual(change('B4000'), ['F!A1']);
  assert.deepEqual(change('C2999'), ['F!A5']);
  assert.deepEqual(change('X11'), ['F!A2']);
  assert.deepEqual(change('X8000'), ['F!A3']);
  assert.deepEqual(change('E5'), ['F!A4']);
  assert.deepEqual(change('B11'), ['F!A1', 'F!A2', 'F!A5']);
  assert.deepEqual(change('A9999'), []);
  assert.deepEqual(['A1', 'A2', 'A3', 'A4', 'A5'].map((cell) => valueAt(workbook, cell, 1)), [2, 2, 1, 1, 4]);
  // Several changes at once.
  const several = engine.recalculate([edit(workbook, 0, 'D4', 5), edit(workbook, 0, 'X8001', 5), edit(workbook, 0, 'A9998', 5)]);
  assert.deepEqual(names(workbook, several), ['F!A3', 'F!A4']);
});

test('shared and array formulas read from a file', async () => {
  const num = (r: string, v: number) => `<c r="${r}"><v>${v}</v></c>`;
  const file = await open(
    await buildXlsx({
      sheets: [
        {
          name: 'S',
          xml: worksheetXml(
            '<sheetData>' +
              `<row r="1">${num('A1', 1)}${num('B1', 10)}<c r="C1"><f t="shared" ref="C1:C3" si="0">A1+B1+$A$1</f><v>12</v></c>` +
              '<c r="E1"><f t="array" ref="E1:E4">A1:A3*B1:B3</f><v>10</v></c><c r="G1"><f t="array" ref="G1">SUM(A1:A3*B1:B3)</f><v>140</v></c></row>' +
              `<row r="2">${num('A2', 2)}${num('B2', 20)}<c r="C2"><f t="shared" si="0"/><v>23</v></c><c r="E2"><v>40</v></c></row>` +
              `<row r="3">${num('A3', 3)}${num('B3', 30)}<c r="C3"><f t="shared" si="0"/><v>34</v></c><c r="E3"><v>90</v></c></row>` +
              '<row r="4"><c r="E4" t="e"><v>#N/A</v></c><c r="F4"><f>E2+E3</f><v>130</v></c>' +
              '<c r="H4"><f t="dataTable" ref="H4:H5" dt2D="0" r1="A1">TABLE(A1)</f><v>77</v></c></row>' +
              '</sheetData>',
          ),
        },
      ],
    }),
  );
  const workbook = await file.workbook();
  const sheet = workbook.sheets[0];
  const engine = engineFor(workbook);
  // What the file stored is what the formulas give: nothing changes.
  assert.deepEqual(engine.recalculate(), []);
  assert.deepEqual(engine.unparsed, []);
  assert.equal(sheet.modified, undefined);

  const set = (address: string, value: number) => {
    const cell = parseCellRef(address) as { row: number; col: number };
    applyOp(workbook, { type: 'setCells', sheet: 0, cells: [{ row: cell.row, col: cell.col, cell: { kind: 'number', value, style: 0 } }] });
    return { sheet: 0, row: cell.row, col: cell.col };
  };
  // A2 feeds the second cell of the shared formula, the array formula, and through its cells F4.
  let result = engine.recalculate([set('A2', 5)]);
  assert.deepEqual(names(workbook, result), ['S!C2', 'S!E2', 'S!F4', 'S!G1']);
  assert.equal(valueAt(workbook, 'C2'), 26);
  assert.equal(valueAt(workbook, 'E2'), 100);
  assert.equal(valueAt(workbook, 'F4'), 190);
  assert.equal(valueAt(workbook, 'G1'), 200);
  // The dependants of a shared formula keep pointing at their master.
  assert.deepEqual(at(workbook, 'C2')?.formula, { text: '', type: 'shared', si: 0 });
  // The cells an array formula fills hold no formula of their own.
  assert.equal(at(workbook, 'E2')?.formula, undefined);
  // Past the end of the result the array formula leaves #N/A.
  assert.deepEqual([at(workbook, 'E4')?.kind, valueAt(workbook, 'E4')], ['error', '#N/A']);

  // $A$1 is in every cell of the shared formula.
  result = engine.recalculate([set('A1', 2)]);
  assert.deepEqual(names(workbook, result), ['S!C1', 'S!C2', 'S!C3', 'S!E1', 'S!G1']);
  assert.deepEqual(['C1', 'C2', 'C3'].map((cell) => valueAt(workbook, cell)), [14, 27, 35]);
  assert.equal(valueAt(workbook, 'E1'), 20);
  // The data table is left alone.
  assert.equal(valueAt(workbook, 'H4'), 77);
  assert.equal(engine.evaluateText('FORMULATEXT(C3)&ISFORMULA(C3)&ISFORMULA(E2)', { sheet: 0, row: 9, col: 9 }), '=A3+B3+$A$1TRUEFALSE');
});

test('array formulas fill their range', () => {
  const workbook = book({
    S: {
      A1: 1,
      A2: 2,
      A3: 3,
      // A scalar result fills every cell; a single row repeats down the rows.
      C1: { f: 'SUM(A1:A3)', type: 'array', ref: 'C1:D2' },
      F1: { f: '{1,2}*A1', type: 'array', ref: 'F1:G2' },
      // The cells do not exist yet in rows 5 and 6.
      I4: { f: 'A1:A3*2', type: 'array', ref: 'I4:I6' },
      // A single cell shows the first element, with no implicit intersection.
      K9: { f: 'A1:A3', type: 'array', ref: 'K9' },
      L1: '=SUM(I4:I6)+D2',
      // A range a damaged file claims: only the formula's own cell is written.
      N1: { f: 'A1+1', type: 'array', ref: 'A1:XFD1048576' },
      O1: { f: 'A1+2', type: 'array', ref: 'nonsense' },
      P1: { f: 'A1+3', type: 'array', ref: 'Q5:Q6' },
    },
  });
  const engine = engineFor(workbook);
  engine.recalculate();
  assert.deepEqual(['C1', 'D1', 'C2', 'D2'].map((cell) => valueAt(workbook, cell)), [6, 6, 6, 6]);
  assert.deepEqual(['F1', 'G1', 'F2', 'G2'].map((cell) => valueAt(workbook, cell)), [1, 2, 1, 2]);
  assert.deepEqual(['I4', 'I5', 'I6'].map((cell) => valueAt(workbook, cell)), [2, 4, 6]);
  assert.equal(valueAt(workbook, 'K9'), 1);
  assert.equal(valueAt(workbook, 'L1'), 18);
  assert.deepEqual(['N1', 'O1', 'P1'].map((cell) => valueAt(workbook, cell)), [2, 3, 4]);
  assert.equal(at(workbook, 'Q5'), undefined);
  assert.equal(workbook.sheets[0].rows.size < 20, true);

  const result = engine.recalculate([edit(workbook, 0, 'A3', 10)]);
  assert.deepEqual(names(workbook, result), ['S!C1', 'S!C2', 'S!D1', 'S!D2', 'S!I6', 'S!L1']);
  assert.equal(valueAt(workbook, 'L1'), 39);
});

test('defined names', () => {
  const workbook = book(
    {
      S: { A1: 1, A2: 2, A3: 3, B1: '=SUM(Total)', B2: '=Rate*100', B3: '=Local', B4: '=Missing+1', B5: '=Twice', B6: '=Loop', B7: '=SUM(Growing)' },
      T: { A1: 50, B3: '=Local', B4: '=S!Local' },
    },
    [
      { name: 'Total', value: 'S!$A$1:$A$3', hidden: false },
      { name: 'Rate', value: '0.25', hidden: false },
      { name: 'Local', value: 'S!$A$1', hidden: false },
      { name: 'Local', value: 'T!$A$1', localSheet: 1, hidden: false },
      { name: 'TWICE', value: 'SUM(total)*2', hidden: false },
      { name: 'Loop', value: 'Loop+1', hidden: false },
      { name: 'Growing', value: 'OFFSET(S!$A$1,0,0,S!$D$1,1)', hidden: false },
      { name: 'Broken', value: 'SUM((', hidden: false },
    ],
  );
  place(workbook.sheets[0], 0, 3, { kind: 'number', value: 2, style: 0 });
  place(workbook.sheets[0], 7, 1, makeCell('=Broken'));
  const engine = engineFor(workbook);
  engine.recalculate();
  assert.deepEqual(['B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8'].map((cell) => valueAt(workbook, cell)), [6, 25, 1, '#NAME?', 12, '#NAME?', 3, '#NAME?']);
  // A name local to a sheet wins on that sheet; with a sheet prefix the name is seen from that sheet.
  assert.deepEqual([valueAt(workbook, 'B3', 1), valueAt(workbook, 'B4', 1)], [50, 1]);

  // The cells behind a name are precedents of the formulas that use it.
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'A2', 20)])), ['S!B1', 'S!B5', 'S!B7']);
  assert.deepEqual([valueAt(workbook, 'B1'), valueAt(workbook, 'B5'), valueAt(workbook, 'B7')], [24, 48, 21]);
  // A name built on OFFSET follows the cell that sizes it.
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'D1', 3)])), ['S!B7']);
  assert.equal(valueAt(workbook, 'B7'), 24);

  // Names are read again after invalidate().
  workbook.definedNames = workbook.definedNames.map((name) => (name.name === 'Rate' ? { ...name, value: '0.5' } : name));
  assert.deepEqual(engine.recalculate(), []);
  engine.invalidate();
  assert.deepEqual(names(workbook, engine.recalculate()), ['S!B2']);
  assert.equal(valueAt(workbook, 'B2'), 50);
});

test('volatile formulas are calculated every time', () => {
  let clock = 45000;
  let dice = 0.1;
  counters.tick = 0;
  const workbook = book({
    S: {
      A1: 1,
      B1: '=NOW()',
      B2: '=RAND()',
      B3: '=B1+1',
      B4: '=NOW()-NOW()',
      C1: '=INDIRECT("A"&A2)',
      C2: '=SUM(OFFSET(A1,0,0,A2,1))',
      C3: '=SUM(A1:INDEX(A1:A9,A2))',
      D1: '=TICK()',
      A2: 1,
      A3: 30,
    },
  });
  const engine = engineFor(workbook, { now: () => clock, random: () => dice });
  engine.recalculate();
  assert.deepEqual(['B1', 'B2', 'B3', 'B4', 'C1', 'C2', 'C3', 'D1'].map((cell) => valueAt(workbook, cell)), [45000, 0.1, 45001, 0, 1, 1, 1, 1]);

  // Nothing they read changed, and still they are calculated; TICK() is not volatile and is not.
  clock = 45002;
  dice = 0.7;
  assert.deepEqual(names(workbook, engine.recalculate([])), ['S!B1', 'S!B2', 'S!B3']);
  assert.deepEqual([valueAt(workbook, 'B1'), valueAt(workbook, 'B2'), valueAt(workbook, 'B3'), valueAt(workbook, 'D1')], [45002, 0.7, 45003, 1]);

  // Formulas that choose their cells while calculating see a change in a cell no reference names.
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'A2', 3)])), ['S!C1', 'S!C2', 'S!C3']);
  assert.deepEqual([valueAt(workbook, 'C1'), valueAt(workbook, 'C2'), valueAt(workbook, 'C3')], [30, 34, 34]);
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'A3', 40)])), ['S!C1', 'S!C2', 'S!C3']);
  assert.deepEqual([valueAt(workbook, 'C1'), valueAt(workbook, 'C2'), valueAt(workbook, 'C3')], [40, 44, 44]);

  // Without options the clock and the dice are the real ones.
  const real = book({ S: { A1: '=NOW()', A2: '=RAND()' } });
  engineFor(real).recalculate();
  const now = valueAt(real, 'A1') as number;
  assert.ok(now > 45000 && now < 80000, String(now));
  const random = valueAt(real, 'A2') as number;
  assert.ok(random >= 0 && random < 1);
  // A source of random numbers that misbehaves is not trusted.
  const bad = book({ S: { A1: '=RAND()' } });
  engineFor(bad, { random: () => NaN, now: () => NaN }).recalculate();
  assert.equal(valueAt(bad, 'A1'), 0);
});

test('circular references give 0 and are reported', () => {
  const workbook = book({
    S: {
      A1: '=A1+1',
      B1: '=B2+1',
      B2: '=B1+1',
      C1: '=C2',
      C2: '=C3',
      C3: '=SUM(C1:C1)+5',
      // Not in a loop, but reads one.
      D1: '=B1+10',
      E1: { f: 'E1+A9', v: 55 },
      F1: '=1+1',
    },
  });
  const engine = engineFor(workbook);
  engine.recalculate();
  for (const cell of ['A1', 'B1', 'B2', 'C1', 'C2', 'C3']) assert.deepEqual([cell, at(workbook, cell)?.kind, valueAt(workbook, cell)], [cell, 'number', 0]);
  // The file has a value for E1 (a workbook with iterative calculation would): nothing E1 reads changed, so it stays.
  assert.equal(valueAt(workbook, 'E1'), 55);
  assert.deepEqual(engine.mismatched, [{ sheet: 0, row: 0, col: 4 }]);
  assert.equal(valueAt(workbook, 'D1'), 10);
  assert.equal(valueAt(workbook, 'F1'), 2);
  const found = engine.circular.map((cell) => cellName(cell.row, cell.col)).sort();
  assert.deepEqual(found, ['A1', 'B1', 'B2', 'C1', 'C2', 'C3', 'E1']);

  // Breaking a loop clears the report for it.
  engine.recalculate([edit(workbook, 0, 'B2', 7)]);
  assert.equal(valueAt(workbook, 'B1'), 8);
  assert.equal(valueAt(workbook, 'D1'), 18);
  assert.deepEqual(engine.circular, []);
  // Making one reports it.
  engine.recalculate([edit(workbook, 0, 'B2', '=D1')]);
  assert.deepEqual(engine.circular.map((cell) => cellName(cell.row, cell.col)).sort(), ['B1', 'B2', 'D1']);
  assert.deepEqual([valueAt(workbook, 'B1'), valueAt(workbook, 'B2'), valueAt(workbook, 'D1')], [0, 0, 0]);
});

/** Column A: A1 = 1, each cell below adds one to the cell above. */
function chainBook(length: number, close = false): Workbook {
  const workbook = book({ S: {} });
  const sheet = workbook.sheets[0];
  place(sheet, 0, 0, close ? makeCell(`=A${length}+1`) : makeCell(1));
  for (let row = 1; row < length; row++) place(sheet, row, 0, makeCell({ f: `A${row}+1`, v: 0 }));
  return workbook;
}

test('a chain of 200,000 cells does not overflow the stack', () => {
  const length = 200000;
  const workbook = chainBook(length);
  const sheet = workbook.sheets[0];
  // The last cell first: the worst order to calculate in.
  const rows = Array.from(sheet.rows.entries()).reverse();
  sheet.rows = new Map(rows);
  // No limit on the clock: what is asked is whether the stack holds, not how fast a machine that
  // is busy with something else gets through the chain. The budget has its own tests.
  const engine = engineFor(workbook, { budget: { milliseconds: 0 } });
  let start = performance.now();
  // The cells hold 0 as the value of the file.
  const result = engine.recalculate(undefined, { overwrite: true });
  const full = performance.now() - start;
  assert.equal(result[0].cells.length, length - 1);
  assert.equal(getCell(sheet, length - 1, 0)?.value, length);
  assert.equal(getCell(sheet, 99999, 0)?.value, 100000);
  assert.deepEqual(engine.circular, []);

  // A change at the head goes through the whole chain.
  start = performance.now();
  engine.recalculate([edit(workbook, 0, 'A1', 11)]);
  const partial = performance.now() - start;
  assert.equal(getCell(sheet, length - 1, 0)?.value, length + 10);
  console.log(`# chain of ${length}: full ${full.toFixed(0)} ms, after a change at the head ${partial.toFixed(0)} ms`);
});

test('a loop of 100,000 cells is found', () => {
  const length = 100000;
  const workbook = chainBook(length, true);
  const engine = engineFor(workbook, { budget: { milliseconds: 0 } });
  engine.recalculate();
  assert.equal(engine.circular.length, length);
  const sheet = workbook.sheets[0];
  for (const row of [0, 1, 4711, length - 1]) assert.equal(getCell(sheet, row, 0)?.value, 0);
});

test('performance: one change in a sheet of 100,000 formulas', () => {
  const rows = 50000;
  const workbook = book({ S: {}, T: {} });
  const sheet = workbook.sheets[0];
  for (let row = 0; row < rows; row++) {
    place(sheet, row, 0, makeCell(row + 1));
    place(sheet, row, 1, makeCell({ f: `A${row + 1}*2`, v: (row + 1) * 2 }));
    place(sheet, row, 2, makeCell({ f: `B${row + 1}+A${row + 1}+$A$1`, v: (row + 1) * 3 + 1 }));
  }
  place(workbook.sheets[1], 0, 0, makeCell({ f: `SUM(S!B1:B${rows})`, v: rows * (rows + 1) }));
  place(workbook.sheets[1], 1, 0, makeCell({ f: 'SUM(S!C1000:C1010)', v: 0 }));
  const engine = engineFor(workbook);

  let start = performance.now();
  // The first call builds the graph; nothing but the changed cell's dependants is calculated.
  let result = engine.recalculate([edit(workbook, 0, 'A1005', 0)]);
  const first = performance.now() - start;
  assert.deepEqual(names(workbook, result), ['S!B1005', 'S!C1005', 'T!A1', 'T!A2']);
  assert.equal(valueAt(workbook, 'A1', 1), rows * (rows + 1) - 2010);

  const times: number[] = [];
  for (let k = 0; k < 20; k++) {
    const address = edit(workbook, 0, `A${2000 + k * 997}`, -k - 1);
    start = performance.now();
    result = engine.recalculate([address]);
    times.push(performance.now() - start);
    assert.equal(result.reduce((count, entry) => count + entry.cells.length, 0), 3);
  }
  times.sort((a, b) => a - b);
  const median = times[10];
  console.log(`# 100,000 formulas: graph + first change ${first.toFixed(0)} ms, later changes median ${median.toFixed(2)} ms, worst ${times[19].toFixed(2)} ms`);
  assert.ok(median < 500, `one change took ${median.toFixed(1)} ms`);

  // The cell every formula of column C reads.
  start = performance.now();
  result = engine.recalculate([edit(workbook, 0, 'A1', 2)]);
  const all = performance.now() - start;
  assert.equal(result[0].cells.length, rows + 1);
  console.log(`# 100,000 formulas: a change read by ${rows} of them ${all.toFixed(0)} ms`);
});

test('formulas that cannot be parsed keep their values', () => {
  const workbook = book({
    S: {
      A1: 1,
      B1: { f: 'SUM(Table1[Amount])+A1', v: 123 },
      B2: { f: 'A1+', v: 'kept' },
      B3: { f: '', v: 9 },
      B4: { f: '', v: 5, type: 'shared', si: 99 },
      C1: { f: 'B1+A1', v: 0 },
    },
  });
  const before = at(workbook, 'B1');
  const engine = engineFor(workbook);
  const result = engine.recalculate(undefined, { overwrite: true });
  // The cells with unknown syntax are not turned into errors, and what reads them uses their values.
  assert.deepEqual(names(workbook, result), ['S!C1']);
  assert.equal(at(workbook, 'B1'), before);
  assert.equal(valueAt(workbook, 'B2'), 'kept');
  assert.equal(valueAt(workbook, 'B3'), 9);
  assert.equal(valueAt(workbook, 'B4'), 5);
  assert.equal(valueAt(workbook, 'C1'), 124);
  assert.deepEqual(engine.unparsed.map((cell) => cellName(cell.row, cell.col)).sort(), ['B1', 'B2', 'B3']);

  // They are not recalculated when what they would read changes either.
  engine.recalculate([edit(workbook, 0, 'A1', 2)]);
  assert.equal(valueAt(workbook, 'B1'), 123);
  assert.equal(valueAt(workbook, 'C1'), 125);

  // Once the formula is one the engine reads, the cell leaves the list.
  engine.recalculate([edit(workbook, 0, 'B2', '=A1+1')]);
  assert.equal(valueAt(workbook, 'B2'), 3);
  assert.deepEqual(engine.unparsed.map((cell) => cellName(cell.row, cell.col)).sort(), ['B1', 'B3']);
  // And the other way round.
  engine.recalculate([edit(workbook, 0, 'B2', { f: 'A1 +* 1', v: 3 })]);
  assert.equal(valueAt(workbook, 'B2'), 3);
  assert.equal(engine.unparsed.length, 3);
});

test('an edited formula changes what the cell depends on', () => {
  const workbook = book({ S: { A1: 1, A2: 2, A3: 3, B1: '=A1', C1: '=B1*2', B2: '=SUM(A1:A3)' } });
  const engine = engineFor(workbook);
  engine.recalculate();
  // B1 now reads A2; C1 follows.
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'B1', '=A2')])), ['S!B1', 'S!C1']);
  assert.equal(valueAt(workbook, 'C1'), 4);
  // A1 is no longer a precedent of B1, A2 is.
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'A1', 100)])), ['S!B2']);
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'A2', 20)])), ['S!B1', 'S!B2', 'S!C1']);
  assert.equal(valueAt(workbook, 'C1'), 40);
  // A range precedent is forgotten as well.
  engine.recalculate([edit(workbook, 0, 'B2', '=A3')]);
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'A1', 1)])), []);
  // A formula replaced by a value; then a new formula in an empty cell.
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'B1', 7)])), ['S!C1']);
  assert.equal(valueAt(workbook, 'C1'), 14);
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'A2', 21)])), []);
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'D9', '=C1+B2')])), ['S!D9']);
  assert.equal(valueAt(workbook, 'D9'), 17);
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'A3', 4)])), ['S!B2', 'S!D9']);

  // A formula changed without telling the engine is not written over, and is picked up afterwards.
  edit(workbook, 0, 'B2', { f: 'A1*1000', v: -1 });
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'A3', 5)])), ['S!D9']);
  assert.equal(valueAt(workbook, 'B2'), -1);
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'A1', 2)])), ['S!B2', 'S!D9']);
  assert.equal(valueAt(workbook, 'B2'), 2000);
});

test('invalidate() after rows are inserted or sheets change', () => {
  const workbook = book({ S: { A1: 1, A2: 2, B1: '=SUM(A1:A2)', C3: '=A2*10' }, T: { A1: '=S!B1' } });
  const engine = engineFor(workbook);
  engine.recalculate();
  assert.equal(valueAt(workbook, 'B1'), 3);

  applyOp(workbook, { type: 'insert', sheet: 0, axis: 'row', at: 1, count: 2 });
  engine.invalidate();
  assert.deepEqual(engine.unparsed, []);
  // The formulas were rewritten by the insert; their values are still right.
  assert.equal(at(workbook, 'B1')?.formula?.text, 'SUM(A1:A4)');
  assert.deepEqual(engine.recalculate(), []);
  // The new rows are inside the range now, and C3 moved to C5 and reads A4.
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'A2', 10)])), ['S!B1', 'T!A1']);
  assert.equal(valueAt(workbook, 'A1', 1), 13);
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'A4', 3)])), ['S!B1', 'S!C5', 'T!A1']);
  assert.equal(valueAt(workbook, 'C5'), 30);

  // A sheet added and one renamed.
  const added = book({ New: { A1: '=Renamed!B1+1' } }).sheets[0];
  workbook.sheets.push(added);
  workbook.sheets[0].name = 'Renamed';
  place(workbook.sheets[1], 0, 0, makeCell('=Renamed!B1'));
  engine.invalidate();
  engine.recalculate();
  assert.equal(valueAt(workbook, 'A1', 2), 15);
  assert.equal(valueAt(workbook, 'A1', 1), 14);
});

test('a formula bar preview stores nothing', () => {
  const workbook = book({ S: { A1: 2, A2: 3, B1: { f: 'A1*100', v: 200 }, B2: { f: 'A2*100', v: -5 } }, T: { A1: 7 } }, [{ name: 'Rate', value: '0.5', hidden: false }]);
  const engine = engineFor(workbook);
  const here = { sheet: 0, row: 4, col: 4 };
  const snapshot = JSON.stringify(Array.from(workbook.sheets[0].rows.entries()));
  assert.equal(engine.evaluateText('=A1+A2', here), 5);
  assert.equal(engine.evaluateText('SUM(A1:A2)*Rate', here), 2.5);
  assert.equal(engine.evaluateText('T!A1&"x"', here), '7x');
  assert.equal(engine.evaluateText('ROW()+COLUMN()', here), 10);
  assert.equal(engine.evaluateText('SHEET()', { sheet: 1, row: 0, col: 0 }), 2);
  assert.equal(engine.evaluateText('A1', { sheet: 1, row: 0, col: 1 }), 7);
  assert.equal(engine.evaluateText('Z99', here), 0);
  assert.equal(engine.evaluateText('A1:A2', { sheet: 0, row: 1, col: 5 }), 3);
  assert.deepEqual(engine.evaluateText('1/0', here), { error: '#DIV/0!' });
  assert.deepEqual(engine.evaluateText('NOSUCH(1)', here), { error: '#NAME?' });
  assert.deepEqual(engine.evaluateText('BOOM()', here), { error: '#VALUE!' });
  // Formulas answer with the values the model holds, right or wrong.
  assert.equal(engine.evaluateText('B1+B2', here), 195);
  // A formula that refers to its own cell just reads what is there.
  assert.equal(engine.evaluateText('A1+1', { sheet: 0, row: 0, col: 0 }), 3);
  assert.equal(engine.evaluateText('1', { sheet: 99, row: -1, col: NaN }), 1);
  assert.throws(() => engine.evaluateText('=A1+', here), (error: unknown) => error instanceof FormulaParseError && error.position === 4);
  assert.throws(() => engine.evaluateText('', here), FormulaParseError);
  assert.equal(JSON.stringify(Array.from(workbook.sheets[0].rows.entries())), snapshot);
  assert.equal(workbook.sheets[0].modified, undefined);
  // And a recalculation afterwards works as before.
  engine.recalculate([edit(workbook, 0, 'A2', 4)]);
  assert.equal(valueAt(workbook, 'B2'), 400);
});

test('nothing a workbook or a caller does makes a recalculation throw', () => {
  const workbook = book({
    S: {
      A1: 1,
      B1: '=BOOM()+A1',
      B2: '=OVERFLOW()',
      B3: '=NOSUCH(A1)',
      B4: '=Missing!A1+A1',
      B5: '=[1]Sheet1!A1+A1',
      B6: '=S:T!A1',
      B7: `=${'('.repeat(250)}A1${')'.repeat(250)}`,
      B8: `=${new Array(2500).fill('A1').join('+')}`,
      B9: { f: 'TABLE(A1)', v: 3, type: 'dataTable', ref: 'B9' },
      C1: '=B1&B2&B3',
      C2: '=DATE1904()',
    },
    T: {},
  });
  const sheet = workbook.sheets[0];
  // Things a damaged file could hold.
  sheet.rows.set(-5, emptyRow());
  sheet.rows.set(1.5, { ...emptyRow(), cells: [makeCell('=A1')] });
  sheet.rows.set(2000000, { ...emptyRow(), cells: [makeCell('=A1')] });
  (sheet.rows.get(0) as Row).cells[20000] = makeCell('=A1');
  const engine = engineFor(workbook);
  engine.recalculate();
  assert.deepEqual(['B1', 'B2', 'B3', 'B4', 'B5', 'B6'].map((cell) => valueAt(workbook, cell)), ['#VALUE!', '#VALUE!', '#NAME?', '#REF!', '#REF!', '#REF!']);
  assert.equal(valueAt(workbook, 'B7'), 1);
  assert.equal(valueAt(workbook, 'B8'), 2500);
  assert.equal(valueAt(workbook, 'B9'), 3);
  assert.equal(valueAt(workbook, 'C1'), '#VALUE!');
  assert.equal(valueAt(workbook, 'C2'), false);

  const junk = [
    { sheet: 9, row: 0, col: 0 },
    { sheet: 0, row: -1, col: 0 },
    { sheet: 0, row: 0.5, col: 0 },
    { sheet: 0, row: 0, col: 99999999 },
    { sheet: 0, row: NaN, col: 0 },
    null,
    undefined,
    {},
    { sheet: 1, row: 500, col: 500 },
  ] as unknown as CellAddress[];
  assert.deepEqual(engine.recalculate(junk), []);
  assert.deepEqual(engine.recalculate('nonsense' as unknown as CellAddress[]), []);
  engine.recalculate([edit(workbook, 0, 'A1', 2)]);
  assert.equal(valueAt(workbook, 'B8'), 5000);

  // An empty workbook, and one with nothing in it at all.
  assert.deepEqual(engineFor(book({})).recalculate(), []);
  assert.deepEqual(engineFor({ sheets: [] } as unknown as Workbook).recalculate([{ sheet: 0, row: 0, col: 0 }]), []);
  assert.equal(engineFor(book({})).evaluateText('1+1', { sheet: 0, row: 0, col: 0 }), 2);
});

test('formulas that wait on many others are calculated once the others are', () => {
  // Every cell of column B reads the whole of column A, which is a chain; column C fans back in.
  const rows = 3000;
  const workbook = book({ S: {} });
  const sheet = workbook.sheets[0];
  place(sheet, 0, 0, makeCell(1));
  for (let row = 1; row < rows; row++) place(sheet, row, 0, makeCell(`=A${row}+1`));
  for (let row = 0; row < 50; row++) place(sheet, row, 1, makeCell(`=SUM(A1:A${rows})+MAX(A:A)`));
  place(sheet, 0, 2, makeCell('=SUM(B1:B50)'));
  place(sheet, 1, 2, makeCell(`=IF(A${rows}>0,INDEX(A:A,${rows}),0)+OFFSET(A1,${rows - 1},0)+INDIRECT("A${rows}")`));
  const engine = engineFor(workbook);
  engine.recalculate();
  const sum = (rows * (rows + 1)) / 2;
  assert.equal(valueAt(workbook, 'B1'), sum + rows);
  assert.equal(valueAt(workbook, 'C1'), 50 * (sum + rows));
  assert.equal(valueAt(workbook, 'C2'), 3 * rows);
  engine.recalculate([edit(workbook, 0, 'A1', 2)]);
  assert.equal(valueAt(workbook, 'C1'), 50 * (sum + rows + rows + 1));
  assert.equal(valueAt(workbook, 'C2'), 3 * (rows + 1));
});

test('deeply nested formulas along a chain of cells', () => {
  // Each cell nests 200 calls around a reference to the cell above: the stack is used up by few of them.
  const rows = 400;
  const workbook = book({ S: { A1: 1 } });
  const sheet = workbook.sheets[0];
  for (let row = 1; row < rows; row++) place(sheet, row, 0, makeCell(`=${'IF(1,SUM('.repeat(100)}A${row}+1${'))'.repeat(100)}`));
  sheet.rows = new Map(Array.from(sheet.rows.entries()).reverse());
  const engine = engineFor(workbook);
  engine.recalculate();
  assert.equal(valueAt(workbook, `A${rows}`), rows);
  engine.recalculate([edit(workbook, 0, 'A1', 5)]);
  assert.equal(valueAt(workbook, `A${rows}`), rows + 4);
});

test('shared formulas built by hand: a cell pushed off the sheet, a group without its master', () => {
  const workbook = book({
    S: {
      A1: 1,
      A2: 2,
      B2: { f: 'A1+A2', type: 'shared', si: 0, ref: 'B1:B2' },
      B1: { f: '', v: 0, type: 'shared', si: 0 },
      C2: { f: '', v: 0, type: 'shared', si: 0 },
      D5: { f: '', v: 42, type: 'shared', si: 7 },
      E5: { f: '', v: 43, type: 'shared' },
    },
  });
  const engine = engineFor(workbook);
  engine.recalculate(undefined, { overwrite: true });
  assert.equal(valueAt(workbook, 'B2'), 3);
  // One row up, "A1" has nowhere to go.
  assert.equal(valueAt(workbook, 'B1'), '#REF!');
  // One column right: B1+B2, and B1 is an error.
  assert.equal(valueAt(workbook, 'C2'), '#REF!');
  assert.equal(valueAt(workbook, 'D5'), 42);
  assert.equal(valueAt(workbook, 'E5'), 43);
  assert.equal(engine.evaluateText('FORMULATEXT(C2)', { sheet: 0, row: 8, col: 8 }), '=B1+B2');
  // C2 is recalculated too, but stays the error it was.
  assert.deepEqual(names(workbook, engine.recalculate([edit(workbook, 0, 'A2', 5)])), ['S!B2']);
  assert.equal(valueAt(workbook, 'B2'), 6);
});

test('the built-in function library is used when none is given', () => {
  const workbook = book({ S: { A1: 1, A2: 2, B1: '=SUM(A1:A2)+IF(A1=1,10,20)' } });
  new CalcEngine(workbook).recalculate();
  assert.equal(valueAt(workbook, 'B1'), 13);
});
