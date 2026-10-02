// Data validation, second round: a value is tried out with everything that
// depends on it calculated, and leaves no trace; rules the engine cannot
// calculate are not applied; limits at their edges; lists and letters.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { LOCALE_EN, LOCALE_TR, type FormatLocale } from '../src/format/locale';
import { DataValidations } from '../src/format/validation';
import * as validationModule from '../src/format/validation';
import { createSheet, emptyStyleSheet, getCell, type Cell, type DataValidation, type Row, type Sheet, type Workbook } from '../src/model/sheet/types';
import { cellFromInput } from '../src/ops/commands';
import { applyOp, type Op } from '../src/ops/ops';

// Added by this round. Where it is missing, a value is tried out the way it was before:
// the cell is put in place and taken back, and nothing is calculated.
const tryOps: <T>(workbook: Workbook, engine: CalcEngine | undefined, ops: Op[], run: () => T) => T =
  (validationModule as unknown as { tryOps?: <T>(workbook: Workbook, engine: CalcEngine | undefined, ops: Op[], run: () => T) => T }).tryOps ||
  ((workbook, _engine, ops, run) => {
    const inverse = ops.map((op) => applyOp(workbook, op));
    try {
      return run();
    } finally {
      for (let i = inverse.length - 1; i >= 0; i--) applyOp(workbook, inverse[i]);
    }
  });

type Value = number | string | boolean | null | { f: string; v?: number | string | boolean };

function sheetOf(grid: { [ref: string]: Value }, name = 'Data', sheetId = 1): Sheet {
  const sheet = createSheet({ name, sheetId, state: 'visible', kind: 'worksheet', part: `/xl/worksheets/sheet${sheetId}.xml` });
  for (const ref of Object.keys(grid)) {
    const { row, col } = at(ref);
    const value = grid[ref];
    let cell: Cell;
    if (value === null) continue;
    if (typeof value === 'number') cell = { kind: 'number', value, style: 0 };
    else if (typeof value === 'string') cell = { kind: 'string', value, style: 0 };
    else if (typeof value === 'boolean') cell = { kind: 'boolean', value, style: 0 };
    else {
      const result = value.v === undefined ? null : value.v;
      cell = { kind: result === null ? 'blank' : typeof result === 'number' ? 'number' : typeof result === 'boolean' ? 'boolean' : 'string', value: result, style: 0, formula: { text: value.f, type: 'normal' } };
    }
    let data = sheet.rows.get(row);
    if (!data) sheet.rows.set(row, (data = { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] }));
    data.cells[col] = cell;
    sheet.rowCount = Math.max(sheet.rowCount, row + 1);
    sheet.colCount = Math.max(sheet.colCount, col + 1);
  }
  return sheet;
}

function at(ref: string): { row: number; col: number } {
  const match = /^([A-Z]+)(\d+)$/.exec(ref) as RegExpExecArray;
  const col = match[1].split('').reduce((value, letter) => value * 26 + letter.charCodeAt(0) - 64, 0) - 1;
  return { row: Number(match[2]) - 1, col };
}

function range(text: string) {
  const [a, b] = text.split(':');
  const first = at(a);
  const last = at(b || a);
  return { top: first.row, left: first.col, bottom: last.row, right: last.col };
}

function rule(sqref: string, type: string, formula1?: string, more: Partial<DataValidation> = {}): DataValidation {
  return {
    ranges: sqref.split(' ').map(range),
    type,
    operator: 'between',
    ...(formula1 === undefined ? {} : { formula1 }),
    allowBlank: false,
    hideDropDown: false,
    showInputMessage: false,
    showErrorMessage: true,
    errorStyle: 'stop',
    ...more,
  };
}

const bookOf = (...sheets: Sheet[]): Workbook => ({ sheets, styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 });

interface Trial {
  wb: Workbook;
  engine: CalcEngine | undefined;
  rules: DataValidations;
  /** Whether text typed into a cell of the first sheet is let through. */
  ok(ref: string, text: string): boolean;
  /** Writes text into a cell of the first sheet for good, and calculates. */
  write(ref: string, text: string): void;
}

function trial(wb: Workbook, options: { locale?: FormatLocale; engine?: boolean } = {}): Trial {
  const engine = options.engine === false ? undefined : new CalcEngine(wb);
  const sheet = wb.sheets[0];
  const rules = new DataValidations(wb, sheet, { engine: () => engine, locale: options.locale });
  const opsFor = (ref: string, text: string) => {
    const { row, col } = at(ref);
    const cell = cellFromInput(wb, sheet, row, col, text, { locale: options.locale });
    return { row, col, cell, ops: [{ type: 'setCells', sheet: 0, cells: [{ row, col, cell }] }] as Op[] };
  };
  return {
    wb,
    engine,
    rules,
    ok(ref, text) {
      const { row, col, cell, ops } = opsFor(ref, text);
      return rules.check(row, col, cell, text, (run) => tryOps(wb, engine, ops, run)).ok;
    },
    write(ref, text) {
      const { row, col, ops } = opsFor(ref, text);
      for (const op of ops) applyOp(wb, op);
      if (engine) engine.recalculate([{ sheet: 0, row, col }]);
    },
  };
}

/** Everything about the sheets that a tried-out value must leave as it was. */
function stateOf(wb: Workbook) {
  return wb.sheets.map((sheet) => ({
    rows: new Map(sheet.rows),
    modified: sheet.modified,
    changedRows: sheet.changedRows,
    changed: sheet.changedRows ? Array.from(sheet.changedRows).sort() : undefined,
    rowCount: sheet.rowCount,
    colCount: sheet.colCount,
  }));
}

function assertUntouched(wb: Workbook, before: ReturnType<typeof stateOf>): void {
  wb.sheets.forEach((sheet, index) => {
    const was = before[index];
    assert.equal(sheet.rows.size, was.rows.size, `rows of sheet ${index}`);
    // The very same row objects: not one cell was replaced, not even by an equal one.
    for (const [key, row] of was.rows) assert.equal(sheet.rows.get(key), row, `row ${key + 1} of sheet ${index}`);
    assert.equal(sheet.modified, was.modified);
    assert.equal(sheet.changedRows, was.changedRows);
    assert.deepEqual(sheet.changedRows ? Array.from(sheet.changedRows).sort() : undefined, was.changed);
    assert.equal(sheet.rowCount, was.rowCount);
    assert.equal(sheet.colCount, was.colCount);
  });
}

test('validation: a custom rule sees the formulas that depend on the value being typed', () => {
  // H2 may hold what keeps H3 (= H2 * 2) at 100 or less.
  const wb = bookOf(sheetOf({ H2: 1, H3: { f: 'H2*2', v: 2 } }));
  wb.sheets[0].dataValidations = [rule('H2', 'custom', '$H$3<=100')];
  const { ok, write, engine } = trial(wb);
  const before = stateOf(wb);
  assert.equal(ok('H2', '60'), false);
  assertUntouched(wb, before);
  assert.equal(ok('H2', '10'), true);
  assert.equal(ok('H2', '50'), true);
  assert.equal(ok('H2', '50.5'), false);
  assert.equal(ok('H2', '51'), false);
  assertUntouched(wb, before);
  assert.equal(getCell(wb.sheets[0], 2, 7)!.value, 2);
  assert.deepEqual(engine!.mismatched, []);
  // A refused value does not lock the cell: what is typed next is judged on its own.
  write('H2', '10');
  assert.equal(getCell(wb.sheets[0], 2, 7)!.value, 20);
  assert.equal(ok('H2', '60'), false);
  assert.equal(ok('H2', '40'), true);
  assert.equal(getCell(wb.sheets[0], 2, 7)!.value, 20);
});

test('validation: the value is tried through a chain of formulas, other sheets included', () => {
  const data = sheetOf({ A1: 1, A2: { f: 'A1+1', v: 2 }, A3: { f: 'A2*Other!B1', v: 20 }, B5: 5 });
  const other = sheetOf({ A1: { f: 'Data!A3+Data!B5', v: 25 }, B1: 10 }, 'Other', 2);
  const wb = bookOf(data, other);
  data.dataValidations = [rule('A1', 'custom', 'Other!$A$1<100'), rule('B5', 'custom', 'AND(Other!$A$1<100,B5>0)')];
  const { ok } = trial(wb);
  const before = stateOf(wb);
  // A1 = 8 gives A2 = 9, A3 = 90, Other!A1 = 95; A1 = 9 gives 105.
  assert.equal(ok('A1', '8'), true);
  assert.equal(ok('A1', '9'), false);
  assert.equal(ok('B5', '79'), true);
  assert.equal(ok('B5', '80'), false);
  assert.equal(ok('B5', '-1'), false);
  assertUntouched(wb, before);
  assert.equal(getCell(other, 0, 0)!.value, 25);
});

test('validation: trying a value out keeps the values of the file and what the engine knows about them', () => {
  // The file says H3 is 999 where the engine calculates 2: the value of the file stands and the cell is listed.
  const wb = bookOf(sheetOf({ H2: 1, H3: { f: 'H2*2', v: 999 }, H4: { f: 'H3+1', v: 1000 } }));
  wb.sheets[0].dataValidations = [rule('H2', 'custom', '$H$3<=100')];
  const { ok, engine } = trial(wb);
  engine!.recalculate();
  assert.deepEqual(engine!.mismatched, [{ sheet: 0, row: 2, col: 7 }]);
  const before = stateOf(wb);
  // Tried with 60 the formula gives 120, with 10 it gives 20: the value of the file is not what is asked.
  assert.equal(ok('H2', '60'), false);
  assert.equal(ok('H2', '10'), true);
  assertUntouched(wb, before);
  assert.equal(getCell(wb.sheets[0], 2, 7)!.value, 999);
  assert.deepEqual(engine!.mismatched, [{ sheet: 0, row: 2, col: 7 }]);
  assert.equal(engine!.incomplete, false);
});

test('validation: a value tried in place of a formula leaves the formula in the graph', () => {
  const wb = bookOf(sheetOf({ G2: 0, H2: { f: 'G2+1', v: 1 }, H3: { f: 'H2*2', v: 2 } }));
  wb.sheets[0].dataValidations = [rule('H2', 'custom', '$H$3<=100')];
  const { ok, write } = trial(wb);
  const before = stateOf(wb);
  assert.equal(ok('H2', '60'), false);
  assert.equal(ok('H2', '7'), true);
  assertUntouched(wb, before);
  assert.equal(getCell(wb.sheets[0], 1, 7)!.formula!.text, 'G2+1');
  // The formula still follows what it reads.
  write('G2', '5');
  assert.equal(getCell(wb.sheets[0], 1, 7)!.value, 6);
  assert.equal(getCell(wb.sheets[0], 2, 7)!.value, 12);
});

test('validation: without an engine a value is still tried in place, and taken back', () => {
  const wb = bookOf(sheetOf({ H2: 1 }));
  wb.sheets[0].dataValidations = [rule('H2', 'custom', 'H2<5'), rule('A1', 'whole', '1', { formula2: '5' })];
  const { ok } = trial(wb, { engine: false });
  const before = stateOf(wb);
  // The formula cannot be calculated: the rule is not applied. Limits written as numbers are.
  assert.equal(ok('H2', '60'), true);
  assert.equal(ok('A1', '6'), false);
  assertUntouched(wb, before);
});

test('validation: a custom rule the engine cannot calculate is not applied; one that gives an error refuses', () => {
  const sheet = sheetOf({ G1: 1, G2: 2, G3: 3 });
  const wb = bookOf(sheet);
  sheet.dataValidations = [
    rule('A1', 'custom', 'NOSUCH(H1)'),
    rule('A2', 'custom', 'REGEXTEST(A2,"^[a-z]+$")'),
    rule('A3', 'custom', 'LET(x,A3,x>0)'),
    rule('A4', 'custom', '[1]Other!$A$1>0'),
    rule('A5', 'custom', 'SUM(First:Last!A1)>0'),
    rule('A6', 'custom', 'NoSuchName>0'),
    rule('A7', 'custom', '(((('),
    // Formulas that can be calculated and give an error: Excel takes an error for FALSE.
    rule('B1', 'custom', '1/0'),
    rule('B2', 'custom', 'VALUE("x")>0'),
    rule('B3', 'custom', 'MATCH(B3,$G$1:$G$3,0)'),
    rule('B4', 'custom', 'NoSuchSheet!A1>0'),
    // Text is not a yes either.
    rule('B5', 'custom', '"yes"'),
  ];
  const { ok } = trial(wb);
  for (const ref of ['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7']) {
    assert.equal(ok(ref, '5'), true, ref);
    assert.equal(ok(ref, 'text'), true, ref);
  }
  assert.equal(ok('B1', '5'), false);
  assert.equal(ok('B2', '5'), false);
  assert.equal(ok('B3', '2'), true);
  assert.equal(ok('B3', '9'), false);
  assert.equal(ok('B4', '5'), false);
  assert.equal(ok('B5', '5'), false);
});

test('validation: limits given by formulas - one that cannot be calculated lifts the rule, an error refuses', () => {
  const sheet = sheetOf({ G1: 10, G2: { f: '1/0', v: '#DIV/0!' } });
  (sheet.rows.get(1) as Row).cells[6] = { kind: 'error', value: '#DIV/0!', style: 0, formula: { text: '1/0', type: 'normal' } };
  const wb = bookOf(sheet);
  sheet.dataValidations = [
    rule('A1', 'whole', 'NOSUCH(1)', { formula2: '5' }),
    rule('A2', 'decimal', '1', { formula2: 'LET(x,1,x)' }),
    rule('A3', 'textLength', 'REGEXTEST("a","a")', { operator: 'lessThan' }),
    rule('A4', 'date', '[1]Other!$A$1', { operator: 'greaterThan' }),
    rule('B1', 'whole', '1/0', { formula2: '5' }),
    rule('B2', 'decimal', '1', { formula2: '$G$2' }),
    rule('B3', 'textLength', 'VALUE("x")', { operator: 'lessThan' }),
    rule('B4', 'whole', '$G$1', { operator: 'lessThanOrEqual' }),
    // A limit that is text, or empty, says nothing that could be compared: as before, nothing is refused.
    rule('C1', 'whole', '"text"', { operator: 'lessThan' }),
  ];
  const { ok } = trial(wb);
  for (const ref of ['A1', 'A2', 'A4']) assert.equal(ok(ref, '1000'), true, ref);
  assert.equal(ok('A3', 'a long text'), true);
  // What can be said without the limits still is: text is not a whole number.
  assert.equal(ok('A1', 'text'), false);
  assert.equal(ok('B1', '3'), false);
  assert.equal(ok('B2', '3'), false);
  assert.equal(ok('B3', 'a'), false);
  assert.equal(ok('B4', '10'), true);
  assert.equal(ok('B4', '11'), false);
  assert.equal(ok('C1', '5'), true);
});

test('validation: a list whose source cannot be calculated lets every value through and shows no button', () => {
  const sheet = sheetOf({ G1: 'a', G2: 'b' });
  const wb = bookOf(sheet);
  sheet.dataValidations = [rule('A1', 'list', 'NOSUCH(G1:G2)'), rule('A2', 'list', 'LET(x,G1:G2,x)'), rule('A3', 'list', 'INDIRECT("nowhere")'), rule('A4', 'list', 'OFFSET($G$1,0,0,2,1)')];
  const { ok, rules } = trial(wb);
  for (const ref of ['A1', 'A2']) {
    assert.equal(ok(ref, 'anything'), true, ref);
    assert.equal(rules.hasList(at(ref).row, 0), false, ref);
    assert.equal(rules.list(at(ref).row, 0), undefined, ref);
  }
  // UNSURE: a source that gives an error (INDIRECT to a name that is not there) is, as in Excel, no list at all.
  assert.equal(ok('A3', 'anything'), true);
  assert.equal(rules.hasList(2, 0), false);
  assert.equal(ok('A4', 'b'), true);
  assert.equal(ok('A4', 'c'), false);
});

test('validation: between and not between include and exclude their limits', () => {
  const sheet = sheetOf({});
  const wb = bookOf(sheet);
  sheet.dataValidations = [
    rule('A1', 'whole', '5', { formula2: '10' }),
    rule('A2', 'whole', '5', { formula2: '10', operator: 'notBetween' }),
    // Limits the wrong way round are taken in order.
    rule('A3', 'whole', '10', { formula2: '5', operator: 'notBetween' }),
    rule('A4', 'decimal', '0.5', { formula2: '1.5', operator: 'notBetween' }),
    rule('A5', 'textLength', '2', { formula2: '3', operator: 'notBetween' }),
    rule('A6', 'whole', '10', { formula2: '5' }),
  ];
  const { ok } = trial(wb);
  const passes = (ref: string, texts: string[]) => texts.map((text) => ok(ref, text));
  assert.deepEqual(passes('A1', ['4', '5', '7', '10', '11']), [false, true, true, true, false]);
  assert.deepEqual(passes('A2', ['4', '5', '7', '10', '11']), [true, false, false, false, true]);
  assert.deepEqual(passes('A3', ['4', '5', '7', '10', '11']), [true, false, false, false, true]);
  assert.deepEqual(passes('A4', ['0.49', '0.5', '1', '1.5', '1.51']), [true, false, false, false, true]);
  assert.deepEqual(passes('A5', ['a', 'ab', 'abc', 'abcd']), [true, false, false, true]);
  assert.deepEqual(passes('A6', ['4', '5', '10', '11']), [false, true, true, false]);
});

test('validation: a list of cells is compared without regard to case, the way the locale pairs its letters', () => {
  const build = (locale: FormatLocale) => {
    const sheet = sheetOf({ G1: 'kıl', G2: 'İstanbul', G3: 'Pear' });
    const wb = bookOf(sheet);
    sheet.dataValidations = [rule('A1', 'list', '$G$1:$G$3')];
    return trial(wb, { locale });
  };
  const turkish = build(LOCALE_TR);
  assert.equal(turkish.ok('A1', 'KIL'), true);
  assert.equal(turkish.ok('A1', 'kil'), false);
  assert.equal(turkish.ok('A1', 'KİL'), false);
  assert.equal(turkish.ok('A1', 'istanbul'), true);
  assert.equal(turkish.ok('A1', 'ISTANBUL'), false);
  assert.equal(turkish.ok('A1', 'PEAR'), true);
  const english = build(LOCALE_EN);
  assert.equal(english.ok('A1', 'KIL'), false);
  assert.equal(english.ok('A1', 'kıl'), true);
  assert.equal(english.ok('A1', 'istanbul'), false);
  assert.equal(english.ok('A1', 'İSTANBUL'), true);
  assert.equal(english.ok('A1', 'pear'), true);
});

test('validation: a named list with empty cells lets every value through where blanks are ignored', () => {
  const sheet = sheetOf({ G1: 'a', G2: 'b', G4: 'd', J1: 'a', J2: 'b' });
  const wb = bookOf(sheet);
  wb.definedNames = [
    { name: 'Gaps', value: 'Data!$G$1:$G$5' },
    { name: 'Full', value: 'Data!$J$1:$J$2' },
  ];
  sheet.dataValidations = [
    rule('A1', 'list', 'Gaps', { allowBlank: true }),
    rule('A2', 'list', 'Gaps'),
    rule('A3', 'list', 'Full', { allowBlank: true }),
    // UNSURE: written as cells, the same source refuses what is not in it, as Excel is described to do.
    rule('A4', 'list', '$G$1:$G$5', { allowBlank: true }),
  ];
  const { ok, rules } = trial(wb);
  assert.equal(ok('A1', 'anything'), true);
  assert.equal(ok('A1', 'a'), true);
  assert.equal(ok('A2', 'anything'), false);
  assert.equal(ok('A2', 'd'), true);
  assert.equal(ok('A3', 'anything'), false);
  assert.equal(ok('A4', 'anything'), false);
  assert.equal(ok('A4', 'd'), true);
  // The list itself is what it was.
  assert.deepEqual(rules.list(0, 0), { items: ['a', 'b', 'd'], truncated: false });
});

test('validation: a value is not tried while the engine is in the middle of a calculation, or cannot finish one', () => {
  const grid: { [ref: string]: Value } = { A1: 1 };
  for (let r = 2; r <= 40; r++) grid[`A${r}`] = { f: `A${r - 1}+1` };
  const wb = bookOf(sheetOf(grid));
  wb.sheets[0].dataValidations = [rule('A1', 'custom', '$A$40<1000')];
  const rules = new DataValidations(wb, wb.sheets[0], { engine: () => engine });
  const cell: Cell = { kind: 'number', value: 5000, style: 0 };
  const ops: Op[] = [{ type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell }] }];
  let ran = 0;
  const tried = (run: () => unknown) =>
    tryOps(wb, engine, ops, () => {
      ran++;
      return run();
    });

  // An engine whose every call ends after one formula: the chain of 39 is not done in a few calls.
  let engine = new CalcEngine(wb, { budget: { work: 1 } });
  engine.recalculate();
  assert.equal(engine.incomplete, true);
  let before = stateOf(wb);
  assert.equal(tryOps(wb, engine, ops, () => ran++), undefined);
  assert.equal(ran, 0);
  assertUntouched(wb, before);
  // The rule is then not applied: the value is let through rather than judged by half a calculation.
  assert.equal(rules.check(0, 0, cell, '5000', tried).ok, true);
  assert.equal(ran, 0);

  // With room to calculate, the same rule refuses the same value.
  engine = new CalcEngine(wb);
  engine.recalculate();
  before = stateOf(wb);
  assert.equal(rules.check(0, 0, cell, '5000', tried).ok, false);
  assert.equal(ran, 1);
  assert.equal(rules.check(0, 0, { kind: 'number', value: 5, style: 0 }, '5', (run) => tryOps(wb, engine, [{ type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell: { kind: 'number', value: 5, style: 0 } }] }], run)).ok, true);
  assertUntouched(wb, before);
});

test('validation: only changes of cells are tried; anything else runs as it is', () => {
  const wb = bookOf(sheetOf({ A1: 1 }));
  const engine = new CalcEngine(wb);
  const before = stateOf(wb);
  assert.equal(tryOps(wb, engine, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }], () => 'ran'), 'ran');
  assert.equal(tryOps(wb, engine, [{ type: 'setCells', sheet: 9, cells: [] }], () => 'ran'), 'ran');
  assert.equal(tryOps(wb, engine, [], () => 'ran'), 'ran');
  assertUntouched(wb, before);
  // A function that throws is passed on, and everything is still put back.
  assert.throws(
    () =>
      tryOps(wb, engine, [{ type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell: { kind: 'number', value: 2, style: 0 } }] }], () => {
        throw new Error('no');
      }),
    /no/,
  );
  assertUntouched(wb, before);
  assert.equal(wb.sheets[0].rows.set, Map.prototype.set);
  assert.equal(Object.keys(wb.sheets[0].rows).length, 0);
});
