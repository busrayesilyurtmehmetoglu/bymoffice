// Tests for the defects the second review of the formula engine confirmed: recalculation in
// slices that never ends, the dependency graph built in one piece, values of the file replaced
// by what a partly supported function gives, formulas over budget left stale without a word,
// formulas a file stores without a value, and a number of smaller ones.
//
// There is no Excel on the machine these tests were written on. Expected values come from
// Excel's documented behaviour; what is not certain is marked UNSURE.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine, type CalcEngineOptions, type CellAddress } from '../src/calc/engine';
import { parseNumberText } from '../src/calc/values';
import { cellName, createSheet, emptyStyleSheet, getCell, parseCellRef, type Cell, type CellFormula, type Row, type Sheet, type Workbook } from '../src/model/sheet/types';

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
    return { kind: kindOf(value, spec.error), value, style: 0, formula };
  }
  return { kind: kindOf(spec), value: spec, style: 0 };
}

const emptyRow = (): Row => ({ customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] });

function place(sheet: Sheet, row: number, col: number, spec: Spec): void {
  let target = sheet.rows.get(row);
  if (!target) sheet.rows.set(row, (target = emptyRow()));
  target.cells[col] = makeCell(spec);
  if (row >= sheet.rowCount) sheet.rowCount = row + 1;
  if (col >= sheet.colCount) sheet.colCount = col + 1;
}

function book(sheets: Record<string, Cells>): Workbook {
  const workbook: Workbook = { sheets: [], styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 };
  Object.keys(sheets).forEach((name, index) => {
    const sheet = createSheet({ name, sheetId: index + 1, state: 'visible', kind: 'worksheet', part: `/xl/worksheets/sheet${index + 1}.xml` });
    for (const address of Object.keys(sheets[name])) {
      const at = parseCellRef(address) as { row: number; col: number };
      place(sheet, at.row, at.col, sheets[name][address]);
    }
    workbook.sheets.push(sheet);
  });
  return workbook;
}

/** Changes a cell the way an edit does: a new cell in a new row object. Returns the address for recalculate(). */
function edit(workbook: Workbook, sheet: number, address: string, spec: Spec): CellAddress {
  const at = parseCellRef(address) as { row: number; col: number };
  const model = workbook.sheets[sheet];
  const row = model.rows.get(at.row);
  const next: Row = row ? { ...row, cells: row.cells.slice() } : emptyRow();
  next.cells[at.col] = makeCell(spec);
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

const NO_LIMIT: CalcEngineOptions = { budget: { milliseconds: 0 } };

/** The value a formula gets in a cell of a sheet holding `cells`. */
function result(formula: string | FormulaSpec, at = 'H2', cells: Cells = {}): Plain | undefined {
  const workbook = book({ S: { ...cells, [at]: typeof formula === 'string' ? `=${formula}` : formula } });
  new CalcEngine(workbook, NO_LIMIT).recalculate();
  return valueAt(workbook, at);
}

function close(actual: unknown, expected: number, label: string): void {
  assert.ok(typeof actual === 'number' && Math.abs(actual - expected) <= 1e-9 * Math.max(1, Math.abs(expected)), `${label}: ${String(actual)} is not ${expected}`);
}

/** A small deterministic generator, so that a failure can be reproduced from its seed. */
function generator(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return state / 4294967296;
  };
}

/** Every formula cell of a workbook with its value, as "sheet!A1". */
function formulaValues(workbook: Workbook): Map<string, Plain> {
  const out = new Map<string, Plain>();
  workbook.sheets.forEach((sheet, index) => {
    sheet.rows.forEach((row, r) => {
      row.cells.forEach((cell, c) => {
        if (cell && cell.formula) out.set(`${index}!${cellName(r, c)}`, cell.value);
      });
    });
  });
  return out;
}

/** Calls recalculate([]) until the engine is done; fails when that takes more than `limit` calls. */
function finish(engine: CalcEngine, limit: number, label: string, each?: () => void): number {
  let calls = 0;
  while (engine.incomplete) {
    engine.recalculate([]);
    calls++;
    if (each) each();
    assert.ok(calls <= limit, `${label}: still incomplete after ${calls} calls`);
  }
  return calls;
}

// ---------------------------------------------------------------------------
// 1. Slices make progress

/** A summary over a column of lookups into a table whose last column holds formulas no lookup reads. */
function lookupBook(rows: number, table: number): { workbook: Workbook; expected: number } {
  const workbook = book({ Ozet: { A1: `=SUM(Veri!D1:D${rows})` }, Veri: {}, Tablo: {} });
  const data = workbook.sheets[1];
  const lookup = workbook.sheets[2];
  for (let j = 1; j <= table; j++) {
    place(lookup, j - 1, 0, j);
    place(lookup, j - 1, 1, j * 10);
    place(lookup, j - 1, 2, j * 2);
    place(lookup, j - 1, 3, `=B${j}*3`);
  }
  let expected = 0;
  for (let i = 1; i <= rows; i++) {
    const key = ((i * 7) % table) + 1;
    place(data, i - 1, 0, key);
    place(data, i - 1, 3, `=VLOOKUP(A${i},Tablo!$A$1:$D$${table},3,FALSE)+SUMIFS(Tablo!$C$1:$C$${table},Tablo!$A$1:$A$${table},A${i})`);
    expected += key * 4;
  }
  return { workbook, expected };
}

test('formulas finished in one slice are not calculated again because of formulas they never read', () => {
  const { workbook, expected } = lookupBook(3000, 200);
  let now = 0;
  const engine = new CalcEngine(workbook, { budget: { milliseconds: 4 }, clock: () => (now += 1) });
  engine.recalculate();
  assert.equal(engine.incomplete, true, 'the budget is too small for one call');
  const calls = finish(engine, 400, 'lookups');
  assert.ok(calls > 3, `${calls} calls`);
  assert.equal(valueAt(workbook, 'A1', 0), expected);
  for (const j of [1, 77, 200]) assert.equal(valueAt(workbook, `D${j}`, 2), j * 30);
  // The same after an edit of the table.
  engine.recalculate([edit(workbook, 2, 'C5', 1000)]);
  finish(engine, 400, 'after an edit');
  let moved = 0;
  for (let i = 1; i <= 3000; i++) if (((i * 7) % 200) + 1 === 5) moved++;
  assert.equal(valueAt(workbook, 'A1', 0), expected + moved * 2 * (1000 - 10));
});

/** A random workbook of two sheets: numbers, sums, ranges, branches never taken, references across sheets. */
function randomCells(seed: number, count: number, cyclic: boolean): Record<string, Cells> {
  const random = generator(seed);
  const name = (k: number) => cellName(Math.floor(k / 4), k % 4);
  const s: Cells = {};
  const t: Cells = {};
  for (let k = 0; k < count; k++) {
    const kind = random();
    if (kind < 0.18) {
      s[name(k)] = 1 + Math.floor(random() * 9);
      continue;
    }
    // Without cycles a formula only reads cells before it.
    const pick = () => name(Math.floor(random() * (cyclic ? count : Math.max(k, 1))) % (cyclic ? count : Math.max(k, 1)));
    const low = () => (cyclic || k > 0 ? pick() : 'Z99');
    let formula: string;
    if (kind < 0.45) formula = `${low()}+${low()}+${1 + Math.floor(random() * 5)}`;
    else if (kind < 0.6) {
      const a = Math.floor(random() * (cyclic ? count : Math.max(k, 1)));
      const b = Math.min(a + Math.floor(random() * 6), (cyclic ? count : Math.max(k, 1)) - 1);
      formula = k === 0 && !cyclic ? '1+1' : `SUM(${name(Math.min(a, b))}:${name(Math.max(a, b))})+1`;
      // A range covers whole rows of the block between its corners; without cycles it must stay before the cell.
      if (!cyclic && Math.floor(Math.max(a, b) / 4) >= Math.floor(k / 4)) formula = `${low()}*2+1`;
    } else if (kind < 0.72) formula = `IF(${low()}>3,${low()},${low()}+2)`;
    else if (kind < 0.84) formula = `IF(FALSE,${low()},${low()}*2)`;
    else if (kind < 0.92) formula = `T!${cellName(Math.floor(random() * 4), 0)}+${low()}`;
    else formula = `COUNT(${low()},${low()})+ROW(${pick()})`;
    s[name(k)] = `=${formula}`;
  }
  for (let r = 0; r < 4; r++) {
    // Without cycles the other sheet only reads numbers.
    t[cellName(r, 0)] = cyclic ? `=S!${name(Math.floor(random() * count))}+${r}` : r % 2 === 0 ? `=B${r + 1}*2` : r + 5;
    t[cellName(r, 1)] = r + 1;
  }
  return { S: s, T: t };
}

test('random workbooks under tiny budgets: done in a number of calls linear in the formulas, same result, no value in between that is neither old nor final', () => {
  for (let seed = 1; seed <= 90; seed++) {
    const cyclic = seed % 3 !== 0;
    const count = 30 + (seed % 5) * 10;
    const sheets = randomCells(seed, count, cyclic);
    const reference = book(sheets);
    const full = new CalcEngine(reference, NO_LIMIT);
    full.recalculate();
    const expected = formulaValues(reference);
    const circular = addresses(full.circular);
    if (!cyclic) assert.deepEqual(circular, [], `seed ${seed}: no cycles were built`);
    const raw = Object.keys(sheets.S).filter((key) => typeof sheets.S[key] === 'number');
    for (const work of [1, 9, 120]) {
      const label = `seed ${seed}, work ${work}`;
      const workbook = book(sheets);
      const engine = new CalcEngine(workbook, { budget: { work, milliseconds: 0 } });
      const formulas = expected.size;
      let before = formulaValues(workbook);
      let target = expected;
      const check = () => {
        formulaValues(workbook).forEach((value, key) => {
          assert.ok(value === before.get(key) || value === target.get(key), `${label}: ${key} is ${String(value)}, was ${String(before.get(key))}, will be ${String(target.get(key))}`);
        });
      };
      engine.recalculate();
      check();
      const found = new Set(addresses(engine.circular));
      finish(engine, 2 * formulas + 10, label, () => {
        check();
        for (const address of addresses(engine.circular)) found.add(address);
      });
      assert.deepEqual(formulaValues(workbook), expected, label);
      assert.deepEqual(Array.from(found).sort(), circular, `${label}: circular cells`);
      assert.deepEqual(engine.overBudget, [], label);

      // An edit of a number, calculated in slices again.
      if (raw.length === 0) continue;
      const address = raw[seed % raw.length];
      const other = book(sheets);
      const whole = new CalcEngine(other, NO_LIMIT);
      whole.recalculate();
      whole.recalculate([edit(other, 0, address, 77)]);
      before = formulaValues(workbook);
      target = formulaValues(other);
      engine.recalculate([edit(workbook, 0, address, 77)]);
      check();
      finish(engine, 2 * formulas + 10, `${label}, after an edit`, check);
      assert.deepEqual(formulaValues(workbook), target, `${label}, after an edit`);
    }
  }
});

test('a cell of a cycle never shows a calculated value between slices', () => {
  // A1 is tied to the cycle only by a branch it does not take.
  const cells: Cells = { A1: '=IF(FALSE,B2,5)', C1: '=A1+1', A2: '=B2+A1', B2: '=A2', C2: '=A2+A1+10' };
  for (const work of [1, 2, 3, 5]) {
    const workbook = book({ S: cells });
    const engine = new CalcEngine(workbook, { budget: { work, milliseconds: 0 } });
    engine.recalculate();
    const seen = () => {
      for (const address of ['A1', 'A2', 'B2']) assert.ok(valueAt(workbook, address) === null || valueAt(workbook, address) === 0, `${address} is ${String(valueAt(workbook, address))} with work ${work}`);
    };
    seen();
    finish(engine, 40, `work ${work}`, seen);
    assert.deepEqual(['A1', 'C1', 'A2', 'B2', 'C2'].map((address) => valueAt(workbook, address)), [0, 1, 0, 0, 10]);
  }
});

test('a call that goes on with what was left does not start the volatile formulas again', () => {
  const cells: Cells = { A1: '=RAND()', Z1: 1 };
  for (let r = 1; r <= 300; r++) cells[`B${r}`] = `=A$1+${r}+SUM(C$1:C$20)`;
  for (let r = 1; r <= 20; r++) cells[`C${r}`] = r;
  const workbook = book({ S: cells });
  let draws = 0;
  const engine = new CalcEngine(workbook, { budget: { work: 200, milliseconds: 0 }, random: () => ++draws / 1000 });
  engine.recalculate([edit(workbook, 0, 'Z1', 2)]);
  assert.equal(engine.incomplete, true);
  const calls = finish(engine, 700, 'volatile');
  assert.ok(calls > 5, `${calls} calls`);
  assert.equal(draws, 1, 'one random number for the whole recalculation');
  const random = valueAt(workbook, 'A1') as number;
  for (let r = 1; r <= 300; r++) assert.equal(valueAt(workbook, `B${r}`), random + r + 210, `B${r}`);
  // The next recalculation draws again.
  engine.recalculate([edit(workbook, 0, 'Z1', 3)]);
  finish(engine, 700, 'volatile again');
  assert.equal(draws, 2);
  assert.equal(valueAt(workbook, 'B300'), 0.002 + 300 + 210);
});

test('a change made between two slices is seen by lookups, and a preview between them disturbs nothing', () => {
  const cells: Cells = {};
  for (let r = 1; r <= 200; r++) {
    cells[`A${r}`] = `k${r}`;
    cells[`B${r}`] = r;
    cells[`D${r}`] = `=VLOOKUP("k${r}",A$1:B$200,2,FALSE)`;
  }
  const workbook = book({ S: cells });
  const engine = new CalcEngine(workbook, { budget: { work: 100, milliseconds: 0 } });
  engine.recalculate();
  assert.equal(engine.incomplete, true);
  assert.equal(valueAt(workbook, 'D200'), null);
  // A preview reads the model as it is: D200 is not calculated for it.
  assert.equal(engine.evaluateText('D200+1', { sheet: 0, row: 0, col: 7 }), 1);
  assert.equal(engine.evaluateText('VLOOKUP("k7",A1:B20,2,FALSE)', { sheet: 0, row: 0, col: 7 }), 7);
  engine.recalculate([]);
  assert.equal(engine.incomplete, true);
  // A value and two keys: the index of the table kept from the first slices is no longer good.
  engine.recalculate([edit(workbook, 0, 'B200', 5000), edit(workbook, 0, 'B1', 1000), edit(workbook, 0, 'A5', 'k150'), edit(workbook, 0, 'A150', 'gone')]);
  finish(engine, 400, 'lookups');
  assert.equal(valueAt(workbook, 'D200'), 5000);
  assert.equal(valueAt(workbook, 'D1'), 1000);
  assert.equal(valueAt(workbook, 'D5'), '#N/A');
  assert.equal(valueAt(workbook, 'D150'), 5);
  for (let r = 2; r < 200; r++) if (r !== 5 && r !== 150) assert.equal(valueAt(workbook, `D${r}`), r, `D${r}`);
});

test('edits made while the graph is still being built are taken in', () => {
  const cells: Cells = {};
  for (let r = 1; r <= 6000; r++) {
    cells[`A${r}`] = r;
    cells[`B${r}`] = `=A${r}*2+SUM(A$1:A$3)`;
  }
  const workbook = book({ S: cells });
  let now = 0;
  const engine = new CalcEngine(workbook, { budget: { milliseconds: 2 }, clock: () => (now += 1) });
  engine.recalculate();
  const edits: [string, Spec][] = [['C1', '=C2+1'], ['C2', '=C1+1'], ['A2', 100], ['B5999', '=A1+0.5'], ['B7', '=B7+1']];
  let calls = 0;
  while (engine.incomplete || edits.length > 0) {
    const next = calls % 2 === 0 ? edits.shift() : undefined;
    engine.recalculate(next ? [edit(workbook, 0, next[0], next[1])] : []);
    assert.ok(++calls < 5000, 'the calls end');
  }
  assert.deepEqual(['C1', 'C2', 'B7', 'B5999', 'B1', 'B6000'].map((address) => valueAt(workbook, address)), [0, 0, 0, 1.5, 2 + 104, 12000 + 104]);
  // Breaking the circle again.
  engine.recalculate([edit(workbook, 0, 'C2', 4)]);
  finish(engine, 5000, 'after breaking the circle');
  assert.equal(valueAt(workbook, 'C1'), 5);
});

// ---------------------------------------------------------------------------
// 2. The dependency graph is built in slices

test('building the graph takes part in the budget: nothing is written before it is done', () => {
  const cells: Cells = {};
  for (let r = 1; r <= 6000; r++) {
    cells[`A${r}`] = r;
    cells[`B${r}`] = `=A${r}*2+SUM(A$1:A$3)`;
  }
  const workbook = book({ S: cells });
  let now = 0;
  const engine = new CalcEngine(workbook, { budget: { milliseconds: 3 }, clock: () => (now += 1) });
  const first = engine.recalculate();
  assert.equal(engine.incomplete, true);
  assert.deepEqual(first, [], 'the first call only builds');
  assert.equal(valueAt(workbook, 'B1'), null);
  let silent = 1;
  let calls = 1;
  while (engine.incomplete) {
    const changed = engine.recalculate([]);
    if (changed.length === 0 && engine.incomplete) silent++;
    assert.ok(++calls < 3000, 'the calls end');
  }
  assert.ok(silent >= 3, `${silent} calls went into building`);
  for (const r of [1, 2999, 6000]) assert.equal(valueAt(workbook, `B${r}`), r * 2 + 6);
  // The same after invalidate(): the graph is built again, in slices again.
  engine.invalidate();
  const again = engine.recalculate([edit(workbook, 0, 'A1', 11)]);
  assert.equal(engine.incomplete, true);
  assert.deepEqual(again, []);
  finish(engine, 3000, 'after invalidate');
  for (const r of [1, 2999, 6000]) assert.equal(valueAt(workbook, `B${r}`), (r === 1 ? 22 : r * 2) + 16);
});

// ---------------------------------------------------------------------------
// 3. Values of the file are not replaced by what the engine only half knows

test('a use of a function the engine does not support keeps the value of the file', () => {
  const frozen: [string, Plain][] = [
    ['CELL("filename",A1)', 'C:\\x\\[b.xlsx]S'],
    ['CELL("format",A1)', 'F2'],
    ['CELL("type",A1)', 'v'],
    ['CELL("width",A1)', 9],
    ['CELL("color",A1)', 0],
    ['CELL("parentheses",A1)', 0],
    ['CELL("prefix",A1)', "'"],
    ['CELL("protect",A1)', 1],
    ['CELL(B1,A1)', 'F2'],
    ['INDIRECT("R1C1",FALSE)', 5],
    ['INDIRECT("R1C1",0)', 5],
    ['INDIRECT("A1",B2)', 5],
    ['INDEX((A1:A3,B1:B3),1,1,2)', 'x'],
    ['INDEX(A1:B3,1,1,B3)', 5],
    ['SUM(INDEX((A1:A3,B1:B3),0,0,2))', 3],
    ['ROWS((A1:A3,B1:B3))', 3],
    ['TEXT(A2,"gg.aa.yyyy")', '15.03.2023'],
    ['TEXT(A2,"aaaa")', 'Mart'],
    ['TEXT(A2,"gggg")', 'Çarşamba'],
    ['TEXT(A2,"TT.MM.JJJJ")', '15.03.2023'],
    ['TEXT(A2,"jj/mm/aaaa")', '15/03/2023'],
  ];
  const cells: Cells = { A1: 5, A2: 45000, A3: 1, B1: 'format', B2: true, B3: 1 };
  frozen.forEach(([formula, value], index) => (cells[cellName(index, 7)] = { f: formula, v: value }));
  // What is supported is calculated as before.
  const live: [string, Plain, Plain][] = [
    ['CELL("row",A2)', 99, 2],
    ['CELL("col",B1)', 99, 2],
    ['CELL("address",A2)', 'x', '$A$2'],
    ['CELL("contents",A1)', 99, 5],
    ['INDIRECT("A1",TRUE)', 99, 5],
    ['INDIRECT("A1",1)', 99, 5],
    ['INDEX(A1:B3,1,1,1)', 99, 5],
    ['SUM((A1:A3,B3:B3))', 99, 45007],
    ['AREAS((A1:A3,B1:B3))', 99, 2],
    ['TEXT(A2,"dd.mm.yyyy")', 'x', '15.03.2023'],
    ['TEXT(A1,"0.00 \\g")', 'x', '5.00 g'],
    ['TEXT(A1,"0 ""gün""")', 'x', '5 gün'],
    ['TEXT(A2,"[$-41F]dd mm")', 'x', '15 03'],
    ['TEXT(A1,"General")', 'x', '5'],
    ['TEXT(0.5,"h:mm AM/PM")', 'x', '12:00 PM'],
    ['TEXT(1234.5,"0.0E+0")', 'x', '1.2E+3'],
  ];
  live.forEach(([formula, value], index) => (cells[cellName(index, 9)] = { f: formula, v: value }));
  const workbook = book({ S: cells });
  const engine = new CalcEngine(workbook, NO_LIMIT);
  // Everything is treated as changed: only what cannot be calculated at all stays.
  engine.recalculate(undefined, { overwrite: true });
  frozen.forEach(([formula, value], index) => assert.equal(valueAt(workbook, cellName(index, 7)), value, formula));
  assert.deepEqual(addresses(engine.unsupported), frozen.map((_entry, index) => `0!${cellName(index, 7)}`).sort());
  live.forEach(([formula, , value], index) => assert.equal(valueAt(workbook, cellName(index, 9)), value, formula));
  // A change of what they read does not touch them either.
  engine.recalculate([edit(workbook, 0, 'A2', 45001), edit(workbook, 0, 'A1', 6)]);
  frozen.forEach(([formula, value], index) => assert.equal(valueAt(workbook, cellName(index, 7)), value, formula));
  assert.equal(valueAt(workbook, cellName(9, 9)), '16.03.2023');
  // Without a value of the file (a formula just typed in) the cell shows what the engine makes of it.
  assert.equal(result('CELL("format",A1)', 'H2', { A1: 1 }), '#VALUE!');
  assert.equal(result('INDIRECT("R1C1",FALSE)', 'H2', { A1: 1 }), '#REF!');
  assert.equal(result('INDEX(A1:B3,1,1,2)', 'H2', { A1: 1 }), '#REF!');
});

test('a formula calculated again although nothing it reads changed keeps a value of the file that differs', () => {
  const cells: Cells = {
    A1: 1234.5,
    A2: 2,
    // A format with the separators of another language: valid here too, with another meaning.
    B1: { f: 'TEXT(A1,"#.##0,00")', v: '1.234,50' },
    // The format is not a constant: nothing can be said before it is calculated.
    B2: { f: 'TEXT(A1,C2)', v: '1.234,50' },
    C2: '#.##0,00',
    B3: { f: 'A1*2', v: 2469 },
    // The last digit differs: the same number.
    B4: { f: 'A2/3', v: 0.6666666666666667 },
    B5: { f: 'A2+1', v: 99 },
    B6: { f: 'B5+1', v: 100 },
    B7: '=A2+1',
    B8: { f: 'B7*2', v: 50 },
    B9: { f: 'OFFSET(A1,1,0)', v: 77 },
    B10: { f: 'INDIRECT("A2")*2', v: 4 },
  };
  const workbook = book({ S: cells });
  const engine = new CalcEngine(workbook, NO_LIMIT);
  const changed = engine.recalculate();
  assert.equal(valueAt(workbook, 'B1'), '1.234,50');
  assert.equal(valueAt(workbook, 'B2'), '1.234,50');
  assert.equal(valueAt(workbook, 'B3'), 2469);
  assert.equal(valueAt(workbook, 'B4'), 0.6666666666666667);
  assert.equal(valueAt(workbook, 'B5'), 99);
  // What reads a value that was kept reads the kept value.
  assert.equal(valueAt(workbook, 'B6'), 100);
  // A formula without a value is calculated, and so is what reads it.
  assert.equal(valueAt(workbook, 'B7'), 3);
  assert.equal(valueAt(workbook, 'B8'), 6);
  assert.equal(valueAt(workbook, 'B9'), 77);
  assert.equal(valueAt(workbook, 'B10'), 4);
  assert.deepEqual(addresses(engine.mismatched), ['0!B1', '0!B2', '0!B5', '0!B9']);
  assert.deepEqual(changed, [{ sheet: 0, cells: [{ row: 6, col: 1 }, { row: 7, col: 1 }] }]);
  assert.equal(workbook.sheets[0].changedRows && workbook.sheets[0].changedRows.has(0), false);

  // The same after invalidate().
  engine.invalidate();
  engine.recalculate();
  assert.equal(valueAt(workbook, 'B1'), '1.234,50');
  assert.equal(valueAt(workbook, 'B5'), 99);
  assert.equal(valueAt(workbook, 'B8'), 6);
  assert.deepEqual(addresses(engine.mismatched), ['0!B1', '0!B2', '0!B5', '0!B9']);

  // The cell is named as changed, but its formula is the one it had (its format was changed, say).
  engine.recalculate([{ sheet: 0, row: 0, col: 1 }, { sheet: 0, row: 4, col: 1 }]);
  assert.equal(valueAt(workbook, 'B1'), '1.234,50');
  assert.equal(valueAt(workbook, 'B5'), 99);

  // A real change of what they read: the value is the engine's from then on.
  engine.recalculate([edit(workbook, 0, 'A2', 5)]);
  assert.equal(valueAt(workbook, 'B5'), 6);
  assert.equal(valueAt(workbook, 'B6'), 7);
  assert.equal(valueAt(workbook, 'B9'), 5);
  assert.equal(valueAt(workbook, 'B10'), 10);
  assert.equal(valueAt(workbook, 'B1'), '1.234,50', 'B1 does not read A2');
  assert.deepEqual(addresses(engine.mismatched), ['0!B1', '0!B2']);
  engine.recalculate([edit(workbook, 0, 'A1', 2)]);
  assert.equal(valueAt(workbook, 'B1'), '2.000');
  assert.equal(valueAt(workbook, 'B3'), 4);
  assert.deepEqual(addresses(engine.mismatched), []);
  // A value the engine wrote is not a value of the file: a full recalculation may replace it.
  engine.recalculate();
  assert.equal(valueAt(workbook, 'B1'), '2.000');
  assert.deepEqual(addresses(engine.mismatched), []);

  // A new formula in a cell with a value.
  const typed = book({ S: { A1: 1, B1: { f: 'A1+1', v: 50 } } });
  const other = new CalcEngine(typed, NO_LIMIT);
  other.recalculate();
  assert.equal(valueAt(typed, 'B1'), 50);
  other.recalculate([edit(typed, 0, 'B1', { f: 'A1+2', v: 50 })]);
  assert.equal(valueAt(typed, 'B1'), 3);

  // Asked for, everything is replaced.
  const forced = book({ S: cells });
  const third = new CalcEngine(forced, NO_LIMIT);
  third.recalculate(undefined, { overwrite: true });
  assert.equal(valueAt(forced, 'B1'), '1234.5000');
  assert.equal(valueAt(forced, 'B5'), 3);
  assert.equal(valueAt(forced, 'B6'), 4);
  assert.deepEqual(addresses(third.mismatched), []);
});

test('formulas that read the clock or a random number are replaced at every recalculation', () => {
  const workbook = book({ S: { A1: { f: 'NOW()', v: 1 }, A2: { f: 'A1+1', v: 2 }, A3: { f: 'RAND()', v: 0.5 }, A4: { f: 'TODAY()', v: 3 }, A5: { f: 'RANDBETWEEN(7,7)', v: 1 } } });
  const engine = new CalcEngine(workbook, { ...NO_LIMIT, now: () => 45000.25, random: () => 0.125 });
  engine.recalculate();
  assert.deepEqual(['A1', 'A2', 'A3', 'A4', 'A5'].map((address) => valueAt(workbook, address)), [45000.25, 45001.25, 0.125, 45000, 7]);
  assert.deepEqual(engine.mismatched, []);
});

test('an array formula whose cells differ from the file keeps all of them', () => {
  const cells: Cells = { A1: 1, A2: 2, B1: { f: 'A1:A2*2', v: 20, type: 'array', ref: 'B1:B2' }, B2: 40, C1: '=B2+1' };
  const workbook = book({ S: cells });
  const engine = new CalcEngine(workbook, NO_LIMIT);
  engine.recalculate();
  assert.deepEqual(['B1', 'B2', 'C1'].map((address) => valueAt(workbook, address)), [20, 40, 41]);
  assert.deepEqual(addresses(engine.mismatched), ['0!B1']);
  engine.recalculate([edit(workbook, 0, 'A2', 3)]);
  assert.deepEqual(['B1', 'B2', 'C1'].map((address) => valueAt(workbook, address)), [2, 6, 7]);
  assert.deepEqual(engine.mismatched, []);
  // The cell of the formula agrees with the file, another cell it fills does not.
  const partly = book({ S: { A1: 1, A2: 2, B1: { f: 'A1:A2*2', v: 2, type: 'array', ref: 'B1:B2' }, B2: 40, C1: '=B2+1' } });
  const other = new CalcEngine(partly, NO_LIMIT);
  other.recalculate();
  assert.deepEqual(['B1', 'B2', 'C1'].map((address) => valueAt(partly, address)), [2, 40, 41]);
  assert.deepEqual(addresses(other.mismatched), ['0!B1']);
  // All of them agree: nothing is written.
  const agreeing = book({ S: { A1: 1, A2: 2, B1: { f: 'A1:A2*2', v: 2, type: 'array', ref: 'B1:B2' }, B2: 4 } });
  const third = new CalcEngine(agreeing, NO_LIMIT);
  assert.deepEqual(third.recalculate(), []);
  assert.deepEqual(third.mismatched, []);
});

test('a formula reached both by a change and by a cell that was only named keeps nothing of the file', () => {
  // B1 reads A1:A2 and D1:D2. A1 is named as changed with the formula it had; D1 reads C1, which did change.
  // Both ways round: the two ranges are looked at in the order the formula names them, or the other.
  for (const formula of ['SUM(A1:A2)+SUM(D1:D2)', 'SUM(D1:D2)+SUM(A1:A2)']) {
    const cells: Cells = { A1: { f: '1+1', v: 2 }, A2: 5, C1: 1, D1: { f: 'C1*1', v: 1 }, D2: 3, B1: { f: formula, v: 999 } };
    const workbook = book({ S: cells });
    const engine = new CalcEngine(workbook, NO_LIMIT);
    engine.recalculate([{ sheet: 0, row: 0, col: 0 }, edit(workbook, 0, 'C1', 10)]);
    assert.equal(valueAt(workbook, 'D1'), 10, formula);
    assert.equal(valueAt(workbook, 'B1'), 20, formula);
    assert.deepEqual(engine.mismatched, [], formula);
  }
});

// ---------------------------------------------------------------------------
// 4. A formula over budget is never passed over in silence

/** E1 scans three columns of `rows` cells; F1 reads E1. */
function heavy(rows: number): Workbook {
  const workbook = book({ S: { E1: { f: `SUMPRODUCT((A1:A${rows}="x")*(B1:B${rows}>5)*C1:C${rows})`, v: -1 }, F1: { f: 'E1*2', v: -2 }, G1: { f: 'B1+1', v: 2 } } });
  const sheet = workbook.sheets[0];
  for (let r = 0; r < rows; r++) {
    place(sheet, r, 0, r % 2 === 0 ? 'x' : 'y');
    place(sheet, r, 1, r % 10);
    place(sheet, r, 2, 1);
  }
  return workbook;
}

test('a formula that needs more than one budget is given a longer one of its own', () => {
  const rows = 20000;
  const workbook = heavy(rows);
  let now = 0;
  // Ten looks at the clock per call; E1 needs far more, but less than the hard limit.
  const engine = new CalcEngine(workbook, { budget: { milliseconds: 10, hardMilliseconds: 5000 }, clock: () => (now += 1) });
  engine.recalculate(undefined, { overwrite: true });
  finish(engine, 20, 'the first calculation');
  // Rows 0, 2, 4, ... hold "x"; of those, the ones whose B is 6 or 8.
  assert.equal(valueAt(workbook, 'E1'), (rows / 10) * 2);
  assert.equal(valueAt(workbook, 'F1'), (rows / 10) * 4);
  assert.deepEqual(engine.overBudget, []);
  assert.deepEqual(engine.stale, []);
  engine.recalculate([edit(workbook, 0, 'B1', 6)]);
  finish(engine, 20, 'after an edit');
  assert.equal(valueAt(workbook, 'E1'), (rows / 10) * 2 + 1);
  assert.equal(valueAt(workbook, 'F1'), (rows / 10) * 4 + 2);
  assert.equal(valueAt(workbook, 'G1'), 7);
  assert.equal(engine.incomplete, false);
});

test('a formula over the hard limit too is reported, with what reads it, and tried once at every change', () => {
  const rows = 20000;
  const workbook = heavy(rows);
  let now = 0;
  let reads = 0;
  const engine = new CalcEngine(workbook, {
    budget: { milliseconds: 5, hardMilliseconds: 15 },
    clock: () => {
      reads++;
      return (now += 1);
    },
  });
  engine.recalculate(undefined, { overwrite: true });
  finish(engine, 20, 'the first calculation');
  assert.equal(valueAt(workbook, 'E1'), -1);
  assert.equal(valueAt(workbook, 'G1'), 1);
  assert.deepEqual(addresses(engine.overBudget), ['0!E1']);
  assert.deepEqual(addresses(engine.stale), ['0!E1', '0!F1']);
  // A change of something else does not try it again.
  reads = 0;
  engine.recalculate([edit(workbook, 0, 'D1', 1)]);
  assert.equal(engine.incomplete, false);
  assert.ok(reads < 5, `${reads} looks at the clock`);
  // A change of what it reads does, once.
  reads = 0;
  engine.recalculate([edit(workbook, 0, 'B1', 6)]);
  const calls = finish(engine, 20, 'after an edit');
  assert.ok(reads >= 15 && reads < 60, `${reads} looks at the clock in ${calls + 1} calls`);
  assert.equal(valueAt(workbook, 'E1'), -1);
  assert.equal(valueAt(workbook, 'G1'), 7);
  assert.deepEqual(addresses(engine.overBudget), ['0!E1']);
  assert.deepEqual(addresses(engine.stale), ['0!E1', '0!F1']);
  reads = 0;
  engine.recalculate([]);
  assert.ok(reads < 5, 'nothing is tried without a change');
  // The same formula once it is cheap: the change that makes it so has it tried again.
  engine.recalculate([edit(workbook, 0, 'E1', { f: `IF(H1=1,0,SUMPRODUCT((A1:A${rows}="x")*(B1:B${rows}>5)*C1:C${rows}))`, v: -1 })]);
  finish(engine, 20, 'guarded');
  assert.deepEqual(addresses(engine.overBudget), ['0!E1']);
  engine.recalculate([edit(workbook, 0, 'H1', 1)]);
  finish(engine, 20, 'cheap now');
  assert.equal(valueAt(workbook, 'E1'), 0);
  assert.equal(valueAt(workbook, 'F1'), 0);
  assert.deepEqual(engine.overBudget, []);
  assert.deepEqual(engine.stale, []);
  // A formula that can be calculated ends the report.
  engine.recalculate([edit(workbook, 0, 'E1', '=COUNT(C1:C100)')]);
  finish(engine, 20, 'a new formula');
  assert.equal(valueAt(workbook, 'E1'), 100);
  assert.equal(valueAt(workbook, 'F1'), 200);
  assert.deepEqual(engine.overBudget, []);
  assert.deepEqual(engine.stale, []);
});

test('invalidate() forgets which formulas were over budget', () => {
  const workbook = heavy(20000);
  const tight = new CalcEngine(workbook, { budget: { work: 1000, hardWork: 3000, milliseconds: 0 } });
  tight.recalculate(undefined, { overwrite: true });
  finish(tight, 20, 'tight');
  assert.deepEqual(addresses(tight.overBudget), ['0!E1']);
  tight.invalidate();
  assert.deepEqual(tight.overBudget, []);
  assert.deepEqual(tight.stale, []);
  tight.recalculate(undefined, { overwrite: true });
  finish(tight, 20, 'tight again');
  assert.deepEqual(addresses(tight.overBudget), ['0!E1'], 'tried again, and given up again');
});

// ---------------------------------------------------------------------------
// 5. SEQUENCE and RANDARRAY with an argument left empty

test('an empty size argument of SEQUENCE and RANDARRAY is 1', () => {
  assert.equal(result('SUM(SEQUENCE(,2))'), 3);
  assert.equal(result('COLUMNS(SEQUENCE(,2))'), 2);
  assert.equal(result('ROWS(SEQUENCE(,2))'), 1);
  assert.equal(result('ROWS(SEQUENCE(2,))'), 2);
  assert.equal(result('COLUMNS(SEQUENCE(2,))'), 1);
  assert.equal(result('SUM(SEQUENCE(2,,5))'), 11);
  assert.equal(result('COLUMNS(RANDARRAY(,2))'), 2);
  assert.equal(result('ROWS(RANDARRAY(,2))'), 1);
  assert.equal(result('ROWS(RANDARRAY(3,))'), 3);
  // A size that is given is still read.
  assert.equal(result('IFERROR(ROWS(SEQUENCE(0,2)),"calc")'), 'calc');
  assert.equal(result('SUM(SEQUENCE(A1,2))', 'H2', { A1: 2 }), 10);
});

// ---------------------------------------------------------------------------
// 6. Formulas a file stores without a value

test('an error the engine wrote itself is not taken for a value of the file later', () => {
  // C3 has a value from before B3 was written without one: it cannot be trusted.
  const cells: Cells = { A1: 1, B1: '=NOSUCH(A1)', B2: '=Missing!A1+1', B3: '=A1+1', B4: '=LET(x,1,x)', B5: { f: 'NOSUCH(A1)', v: 7 }, C3: { f: 'B3*10', v: 5 } };
  const workbook = book({ S: cells });
  const engine = new CalcEngine(workbook, NO_LIMIT);
  assert.deepEqual(addresses(engine.uncalculated()), ['0!B1', '0!B2', '0!B3', '0!B4']);
  engine.recalculate(engine.uncalculated());
  assert.deepEqual(['B1', 'B2', 'B3', 'B4', 'B5'].map((address) => valueAt(workbook, address)), ['#NAME?', '#REF!', 2, '#NAME?', 7]);
  assert.equal(valueAt(workbook, 'C3'), 20);
  assert.deepEqual(addresses(engine.unsupported), ['0!B5']);
  assert.deepEqual(engine.uncalculated(), []);
  engine.invalidate();
  engine.recalculate();
  assert.deepEqual(addresses(engine.unsupported), ['0!B5']);
  assert.deepEqual(['B1', 'B2', 'B3', 'B4', 'B5'].map((address) => valueAt(workbook, address)), ['#NAME?', '#REF!', 2, '#NAME?', 7]);
  // Still calculated: the function arrives (here: the formula changes to one that exists).
  engine.recalculate([edit(workbook, 0, 'A1', 5)]);
  assert.equal(valueAt(workbook, 'B3'), 6);
  assert.deepEqual(addresses(engine.unsupported), ['0!B5']);
});

// ---------------------------------------------------------------------------
// 7. Smaller ones

test('CUMIPMT and CUMPRINC refuse periods outside the loan', () => {
  for (const name of ['CUMIPMT', 'CUMPRINC']) {
    assert.equal(result(`${name}(0.01,12,1000,1,13,0)`), '#NUM!', `${name}: end past the last period`);
    assert.equal(result(`${name}(0.01,12,1000,0,5,0)`), '#NUM!', `${name}: start before the first`);
    assert.equal(result(`${name}(0.01,12,1000,5,4,0)`), '#NUM!', `${name}: start after end`);
    assert.equal(typeof result(`${name}(0.01,12,1000,1,12,0)`), 'number', name);
  }
  close(result('CUMPRINC(0.01,12,1000,1,12,0)'), -1000, 'the whole loan');
});

test('a calculated criterion of a database function that is a constant TRUE or FALSE', () => {
  const cells: Cells = { A1: 'Name', B1: 'Qty', A2: 'a', B2: 1, A3: 'b', B3: 2, A4: 'c', B4: 4, D1: 'Any', D2: true, E1: null, E2: true, F1: 'Any', F2: false, G1: 'Any', G2: 5, I1: 'Name', J1: 'Any', I2: 'b', J2: true, I3: 'c', J3: false };
  assert.equal(result('DSUM(A1:B4,"Qty",D1:D2)', 'H9', cells), 7);
  assert.equal(result('DSUM(A1:B4,"Qty",E1:E2)', 'H9', cells), 7);
  assert.equal(result('DCOUNT(A1:B4,"Qty",F1:F2)', 'H9', cells), 0);
  assert.equal(result('DSUM(A1:B4,"Qty",F1:F2)', 'H9', cells), 0);
  // TRUE asks for nothing more; a row with FALSE finds nothing.
  assert.equal(result('DSUM(A1:B4,"Qty",I1:J3)', 'H9', cells), 2);
  // Anything else under a heading that names no field still cannot be applied.
  assert.equal(result('DSUM(A1:B4,"Qty",G1:G2)', 'H9', cells), '#VALUE!');
});

test('text written as a whole number and a fraction is a number', () => {
  assert.equal(result('"0 1/2"+0'), 0.5);
  assert.equal(result('"1 1/2"+0'), 1.5);
  assert.equal(result('"-2 1/4"+0'), -2.25);
  assert.equal(result('VALUE(" 3 3/4 ")'), 3.75);
  assert.equal(result('"1 1/0"+0'), '#VALUE!');
  assert.equal(result('"1 /2"+0'), '#VALUE!');
  assert.equal(parseNumberText('10 10/4'), 12.5);
  // Without the whole number it is a date, as before. UNSURE about the year: the one of the clock.
  const half = parseNumberText('1/2');
  assert.ok(typeof half === 'number' && half > 40000, String(half));
});

test('XNPV and XIRR over an array of rates read their ranges once', () => {
  const cells: Cells = {};
  for (let r = 1; r <= 400; r++) {
    cells[`A${r}`] = r === 1 ? -1000 : 10;
    cells[`B${r}`] = 45000 + r * 30;
    cells[`C${r}`] = 0.01 + r / 10000;
  }
  const workbook = book({ S: cells });
  // 400 rates times 2 x 400 cells is 320,000 cells when the ranges are read again for every rate.
  const engine = new CalcEngine(workbook, { budget: { work: 30000, hardWork: 30000, milliseconds: 0 } });
  const total = engine.evaluateText('SUMPRODUCT(XNPV(C1:C400,A1:A400,B1:B400))', { sheet: 0, row: 0, col: 7 });
  assert.equal(typeof total, 'number', JSON.stringify(total));
  const free = new CalcEngine(workbook, NO_LIMIT);
  const one = free.evaluateText('XNPV(C7,A1:A400,B1:B400)', { sheet: 0, row: 0, col: 7 });
  const seventh = free.evaluateText('INDEX(XNPV(C1:C400,A1:A400,B1:B400),7)', { sheet: 0, row: 0, col: 7 });
  assert.equal(seventh, one);
  const rate = engine.evaluateText('SUMPRODUCT(XIRR(A1:A400,B1:B400,C1:C40))', { sheet: 0, row: 0, col: 7 });
  assert.equal(typeof rate, 'number', JSON.stringify(rate));
});

test('batch(): previews in a row share one memo and one budget', () => {
  const cells: Cells = {};
  for (let r = 1; r <= 2000; r++) {
    cells[`A${r}`] = `k${r}`;
    cells[`B${r}`] = r;
  }
  const workbook = book({ S: cells });
  const at = { sheet: 0, row: 0, col: 5 };
  // Each lookup scans half the table when no index is kept between the calls.
  const engine = new CalcEngine(workbook, { budget: { work: 40000, hardWork: 40000, milliseconds: 0 } });
  const values = engine.batch(() => {
    const out: unknown[] = [];
    for (let k = 1; k <= 300; k++) out.push(engine.evaluateText(`VLOOKUP("k${k * 6}",A1:B2000,2,FALSE)`, at));
    return out;
  });
  assert.deepEqual(values, Array.from({ length: 300 }, (_unused, k) => (k + 1) * 6));
  // One call after the batch has a budget of its own again, and sees a change of the model.
  edit(workbook, 0, 'B6', 66);
  assert.equal(engine.evaluateText('VLOOKUP("k6",A1:B2000,2,FALSE)', at), 66);
  // The budget is one for the whole batch.
  const spent = engine.batch(() => [engine.evaluateText('SUM(B1:B2000)', at), engine.evaluateText('SUMPRODUCT(B1:B2000*B1:B2000*B1:B2000*B1:B2000*B1:B2000*B1:B2000*B1:B2000*B1:B2000*B1:B2000*B1:B2000)', at), engine.evaluateText('SUM(B1:B2000)', at)]);
  assert.equal(typeof spent[0], 'number');
  assert.deepEqual(spent[2], { error: '#VALUE!' });
  assert.equal(typeof engine.evaluateText('SUM(B1:B2000)', at), 'number');
});

test('evaluateShifted(): a formula written for one cell, calculated for another', () => {
  const workbook = book({ S: { A1: 1, A2: 2, A3: 3, B1: 10, B2: 20, B3: 30 }, T: { A1: 100, A2: 200 } });
  const engine = new CalcEngine(workbook, NO_LIMIT);
  const anchor = { sheet: 0, row: 0, col: 2 };
  assert.equal(engine.evaluateShifted('A1+$B$1', anchor, { row: 0, col: 2 }), 11);
  assert.equal(engine.evaluateShifted('A1+$B$1', anchor, { row: 1, col: 2 }), 12);
  assert.equal(engine.evaluateShifted('A1+$B$1', anchor, { row: 2, col: 3 }), 40);
  assert.equal(engine.evaluateShifted('$A1+B$1', anchor, { row: 2, col: 3 }), 3);
  assert.equal(engine.evaluateShifted('SUM(A$1:A1)', anchor, { row: 2, col: 2 }), 6);
  assert.equal(engine.evaluateShifted('T!A1*2', anchor, { row: 1, col: 2 }), 400);
  assert.equal(engine.evaluateShifted('ROW()*10+COLUMN()', anchor, { row: 4, col: 5 }), 56);
  // A reference pushed off the sheet.
  assert.deepEqual(engine.evaluateShifted('A1', { sheet: 0, row: 5, col: 5 }, { row: 4, col: 5 }), { error: '#REF!' });
  assert.deepEqual(engine.evaluateShifted('A1', { sheet: 0, row: 5, col: 5 }, { row: 5, col: 4 }), { error: '#REF!' });
  assert.equal(engine.evaluateShifted('$A$1', { sheet: 0, row: 5, col: 5 }, { row: 0, col: 0 }), 1);
  assert.throws(() => engine.evaluateShifted('1+', anchor, { row: 0, col: 0 }));
  // The same as writing the shifted formula out.
  assert.equal(engine.evaluateShifted('A1*B1', anchor, { row: 2, col: 2 }), engine.evaluateText('A3*B3', { sheet: 0, row: 2, col: 2 }));
});

test('circular references follow the references a formula names, not the ones it reads', () => {
  // Asking where a cell is does not read it.
  const where = book({ S: { A1: '=CELL("row",A1)', A2: '=CELL("col",A2)+CELL("address",A2)&""', A3: '=CELL("contents",A3)', B1: '=CELL("row",C1)', C1: '=B1+1' } });
  const first = new CalcEngine(where, NO_LIMIT);
  first.recalculate();
  assert.equal(valueAt(where, 'A1'), 1);
  assert.equal(valueAt(where, 'A3'), 0);
  assert.equal(valueAt(where, 'B1'), 1);
  assert.equal(valueAt(where, 'C1'), 2);
  assert.deepEqual(addresses(first.circular), ['0!A3']);
  // A branch that is never calculated ties its cells together all the same, as in Excel.
  const branch = book({ S: { A1: '=IF(TRUE,1,A1)', B1: '=IF(FALSE,C1,1)', C1: '=B1', D1: '=C1+5', E1: '=IF(FALSE,F1,1)', F1: 3 } });
  const second = new CalcEngine(branch, NO_LIMIT);
  second.recalculate();
  assert.deepEqual(['A1', 'B1', 'C1', 'D1', 'E1'].map((address) => valueAt(branch, address)), [0, 0, 0, 5, 1]);
  assert.deepEqual(addresses(second.circular), ['0!A1', '0!B1', '0!C1']);
  // Breaking the tie.
  second.recalculate([edit(branch, 0, 'C1', 4)]);
  assert.deepEqual(['B1', 'C1', 'D1'].map((address) => valueAt(branch, address)), [1, 4, 9]);
  assert.deepEqual(second.circular, []);
  // Making one through an edit; only the cells of the cycle are 0.
  second.recalculate([edit(branch, 0, 'F1', '=E1+D1')]);
  assert.deepEqual(['E1', 'F1', 'D1'].map((address) => valueAt(branch, address)), [0, 0, 9]);
  assert.deepEqual(addresses(second.circular), ['0!E1', '0!F1']);
  second.recalculate([edit(branch, 0, 'E1', '=2')]);
  assert.deepEqual(['E1', 'F1'].map((address) => valueAt(branch, address)), [2, 11]);
  assert.deepEqual(second.circular, []);
});

test('a circle through cells chosen while calculating ends when the cells change', () => {
  const workbook = book({ S: { A1: '=OFFSET(A1,1,0)+1', A2: '=INDIRECT("A1")+1', A3: '=A2+1' } });
  const engine = new CalcEngine(workbook, NO_LIMIT);
  engine.recalculate();
  assert.deepEqual(['A1', 'A2', 'A3'].map((address) => valueAt(workbook, address)), [0, 0, 1]);
  assert.deepEqual(addresses(engine.circular), ['0!A1', '0!A2']);
  engine.recalculate([edit(workbook, 0, 'A2', 5)]);
  assert.deepEqual(['A1', 'A2', 'A3'].map((address) => valueAt(workbook, address)), [6, 5, 6]);
  assert.deepEqual(engine.circular, []);
  engine.recalculate([edit(workbook, 0, 'A2', '=INDIRECT("A1")+1')]);
  assert.deepEqual(['A1', 'A2', 'A3'].map((address) => valueAt(workbook, address)), [0, 0, 1]);
  assert.deepEqual(addresses(engine.circular), ['0!A1', '0!A2']);
});

test('a reference that is only asked where it is ties nothing into a cycle', () => {
  const workbook = book({ S: { A1: '=ROW(B1)+COLUMNS(B1:C1)', B1: '=A1+1', C1: '=OFFSET(D1,0,0)*0+B1', D1: '=C1' } });
  const engine = new CalcEngine(workbook, NO_LIMIT);
  engine.recalculate();
  assert.equal(valueAt(workbook, 'A1'), 3);
  assert.equal(valueAt(workbook, 'B1'), 4);
  // C1 reads D1 through OFFSET, D1 reads C1: found while calculating.
  assert.deepEqual(addresses(engine.circular), ['0!C1', '0!D1']);
});

test('cycles in large sheets are found without comparing every formula with every other', () => {
  // 20,000 formulas that each read the whole column they are in, and one another through it.
  const workbook = book({ S: {} });
  const sheet = workbook.sheets[0];
  const rows = 20000;
  for (let r = 0; r < rows; r++) {
    place(sheet, r, 0, r);
    place(sheet, r, 1, `=A${r + 1}/SUM(C$1:C$${rows})`);
    place(sheet, r, 2, `=B${r + 1}+1`);
  }
  place(sheet, 0, 4, `=SUM(A1:A${rows})`);
  const engine = new CalcEngine(workbook, { budget: { milliseconds: 0 } });
  const start = Date.now();
  engine.recalculate();
  const took = Date.now() - start;
  assert.equal(engine.circular.length, 2 * rows);
  assert.equal(valueAt(workbook, 'B7'), 0);
  assert.equal(valueAt(workbook, 'E1'), (rows * (rows - 1)) / 2);
  assert.ok(took < 50000, `took ${took} ms`);
});

// ---------------------------------------------------------------------------
// 8. Mutations that survived

test('implicit intersection: a formula above or below a column, left or right of a row', () => {
  const cells: Cells = { A3: 1, A4: 2, A5: 3, C7: 10, D7: 20 };
  assert.equal(result('A3:A5+0', 'B2', cells), '#VALUE!');
  assert.equal(result('A3:A5+0', 'B3', cells), 1);
  assert.equal(result('A3:A5+0', 'B5', cells), 3);
  assert.equal(result('A3:A5+0', 'B6', cells), '#VALUE!');
  assert.equal(result('C7:D7+0', 'B9', cells), '#VALUE!');
  assert.equal(result('C7:D7+0', 'C9', cells), 10);
  assert.equal(result('C7:D7+0', 'D9', cells), 20);
  assert.equal(result('C7:D7+0', 'E9', cells), '#VALUE!');
});

test('text as a number: the length limit, and text that only the cell editor gives a meaning', () => {
  // 64 characters are still looked at, 65 are not.
  const money = (length: number) => '$' + '1'.padStart(length - 1, ' ');
  assert.equal(money(64).length, 64);
  assert.equal(parseNumberText(money(64)), 1);
  assert.equal(parseNumberText(money(65)), undefined);
  assert.equal(parseNumberText('1,000'.padStart(64, ' ')), 1000);
  assert.equal(parseNumberText('1,000'.padStart(65, ' ')), undefined);
  assert.equal(parseNumberText('=1'), undefined);
  assert.equal(parseNumberText('=1+1'), undefined);
  assert.equal(parseNumberText("'1"), undefined);
  assert.equal(parseNumberText("'$5"), undefined);
  assert.equal(parseNumberText(''), undefined);
  assert.equal(result('"=1"+0'), '#VALUE!');
  assert.equal(result('"\'$5"+0'), '#VALUE!');
});

test('a lookup table read while a cycle is being found gives what a table read afterwards gives', () => {
  // The index of the table is built by the first lookup, while A3 (a key) is still being found circular.
  const cells: Cells = { A1: 0, A2: 5, A3: '=D1*0+A4', A4: '=A3', B1: 'zero', B2: 'five', B3: 'loop', B4: 'loop2' };
  for (let k = 1; k <= 6; k++) cells[`D${k}`] = `=VLOOKUP(${k % 2 === 0 ? 5 : 0},A$1:B$4,2,FALSE)`;
  const workbook = book({ S: cells });
  const engine = new CalcEngine(workbook, NO_LIMIT);
  engine.recalculate();
  // D1 belongs to no cycle: it reads the table, and the table reads it only through A3, which is 0 in the end.
  assert.deepEqual(addresses(engine.circular), ['0!A3', '0!A4', '0!D1']);
  assert.deepEqual([2, 3, 4, 5, 6].map((k) => valueAt(workbook, `D${k}`)), ['five', 'zero', 'five', 'zero', 'five']);
});

test('exact lookups tell a number from text that looks like it', () => {
  const cells: Cells = { A1: '5', A2: 5, A3: '05', A4: 5.0, B1: 'text', B2: 'number', B3: 'padded', B4: 'again' };
  const formulas: [string, Plain][] = [
    ['VLOOKUP(5,A1:B4,2,FALSE)', 'number'],
    ['VLOOKUP("5",A1:B4,2,FALSE)', 'text'],
    ['VLOOKUP("05",A1:B4,2,FALSE)', 'padded'],
    ['MATCH(5,A1:A4,0)', 2],
    ['MATCH("5",A1:A4,0)', 1],
    ['XLOOKUP(5,A1:A4,B1:B4,"none",0,-1)', 'again'],
    ['XLOOKUP("5",A1:A4,B1:B4,"none",0,-1)', 'text'],
  ];
  // Three times over, so that the index is used.
  const all: Cells = { ...cells };
  for (let round = 0; round < 3; round++) formulas.forEach(([formula], index) => (all[cellName(index + round * 10, 5)] = `=${formula}`));
  const workbook = book({ S: all });
  new CalcEngine(workbook, NO_LIMIT).recalculate();
  for (let round = 0; round < 3; round++) formulas.forEach(([formula, expected], index) => assert.equal(valueAt(workbook, cellName(index + round * 10, 5)), expected, `${formula} (${round})`));
});
