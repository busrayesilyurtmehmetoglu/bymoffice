import assert from 'node:assert/strict';
import { test } from 'node:test';
import { FUNCTIONS } from '../src/calc/functions/index';
import { calcError, isArray, isError, type CalcContext, type ErrorCode, type Scalar, type Value } from '../src/calc/values';
import { MAX_COLS, MAX_ROWS, parseRange } from '../src/model/sheet/types';

// ---------------------------------------------------------------------------
// A sheet made of a small map of cells, enough of a CalcContext to call functions directly.

type Cells = { [name: string]: Scalar };

/** Serial of 15 March 2024 12:00, the clock of every test context. */
const NOW = 45366.5;

function makeContext(cells: Cells, date1904 = false): CalcContext {
  const values = new Map<number, Scalar>();
  let rows = 0;
  let cols = 0;
  for (const name of Object.keys(cells)) {
    const at = parseRange(name)!;
    values.set(at.top * MAX_COLS + at.left, cells[name]);
    rows = Math.max(rows, at.top + 1);
    cols = Math.max(cols, at.left + 1);
  }
  const keys = Array.from(values.keys()).sort((a, b) => a - b);
  let seed = 12345;
  return {
    cellValue: (_sheet, row, col) => (values.has(row * MAX_COLS + col) ? values.get(row * MAX_COLS + col)! : null),
    forEachCell(_sheet, range, visit) {
      for (const key of keys) {
        const row = Math.floor(key / MAX_COLS);
        const col = key % MAX_COLS;
        if (row >= range.top && row <= range.bottom && col >= range.left && col <= range.right) {
          visit(values.get(key)!, row, col);
        }
      }
    },
    usedExtent: () => ({ rows, cols }),
    sheetIndex: () => 0,
    sheetName: () => 'Sheet1',
    sheetCount: () => 1,
    sheet: 0,
    row: 0,
    col: 0,
    date1904,
    now: () => NOW,
    random() {
      seed = (seed * 1103515245 + 12345) % 2147483648;
      return seed / 2147483648;
    },
  };
}

let context = makeContext({});

/** A reference: "A1", "A1:B3", or a whole column such as "A:A" or "A:C". */
function ref(name: string): Value {
  const columns = /^([A-Z]+):([A-Z]+)$/.exec(name);
  if (columns) {
    const left = parseRange(columns[1] + '1')!.left;
    const right = parseRange(columns[2] + '1')!.left;
    return { ref: true, sheet: 0, range: { top: 0, left, bottom: MAX_ROWS - 1, right } };
  }
  return { ref: true, sheet: 0, range: parseRange(name)! };
}

/** An array constant from its rows. */
function arr(rows: Scalar[][]): Value {
  const data: Scalar[] = [];
  for (const row of rows) for (const value of row) data.push(value);
  return { rows: rows.length, cols: rows[0].length, data };
}

function err(code: ErrorCode): Scalar {
  return calcError(code);
}

function call(name: string, ...args: Value[]): Value {
  const spec = FUNCTIONS.get(name);
  assert.ok(spec, `${name} is registered`);
  assert.ok(args.length >= spec.min && (spec.max === undefined || args.length <= spec.max), `${name} takes ${args.length} arguments`);
  return spec.call(args, context);
}

/** A result in a form that is easy to compare: errors as their code, arrays as rows. */
function plain(value: Value): unknown {
  if (isError(value)) return value.error;
  if (isArray(value)) {
    const rows: unknown[][] = [];
    for (let r = 0; r < value.rows; r++) {
      const row: unknown[] = [];
      for (let c = 0; c < value.cols; c++) row.push(plain(value.data[r * value.cols + c]));
      rows.push(row);
    }
    return rows;
  }
  return value;
}

/** Checks rows of [expected, function, ...arguments]; numbers are compared to 12 significant digits. */
function check(rows: [unknown, string, ...Value[]][]): void {
  for (const [expected, name, ...args] of rows) {
    const actual = plain(call(name, ...args));
    const label = `${name}(${args.map((arg) => JSON.stringify(plain(arg))).join(', ')})`;
    if (typeof expected === 'number' && typeof actual === 'number') {
      const scale = Math.max(1, Math.abs(expected));
      assert.ok(Math.abs(actual - expected) <= 1e-12 * scale, `${label}: expected ${expected}, got ${actual}`);
    } else assert.deepEqual(actual, expected, label);
  }
}

test('PMT, FV, PV and NPER', () => {
  check([
    // Values from Excel's documentation, to the cent, and exact formulas.
    [-1037.0320893591636, 'PMT', 0.08 / 12, 10, 10000],
    [-1030.1643271779772, 'PMT', 0.08 / 12, 10, 10000, 0, 1],
    [-1342.0540575303496, 'PMT', 0.05 / 12, 360, 250000],
    [-1000, 'PMT', 0, 10, 10000],
    [-1100, 'PMT', 0, 10, 10000, 1000],
    ['#NUM!', 'PMT', 0.1, 0, 1000],
    ['#VALUE!', 'PMT', 'x', 10, 1000],
    [2581.4033740601362, 'FV', 0.06 / 12, 10, -200, -500, 1],
    [2000, 'FV', 0, 10, -200],
    [-1100, 'FV', 0.1, 1, 0, 1000],
    [-59777.14585118777, 'PV', 0.08 / 12, 240, 500, 0, 0],
    [-5000, 'PV', 0, 10, 500],
    [-1000, 'PV', 0.1, 1, 0, 1100],
    [59.67386567429457, 'NPER', 0.12 / 12, -100, -1000, 10000, 1],
    [10, 'NPER', 0, -100, 1000],
    ['#DIV/0!', 'NPER', 0, 0, 1000],
    [-9.578594039813161, 'NPER', 0.01, 100, 1000],
    // The payment does not even cover the interest.
    ['#NUM!', 'NPER', 0.1, -50, 1000],
  ]);
  const cents = (value: Value): number => Math.round((value as number) * 100) / 100;
  assert.equal(cents(call('PMT', 0.08 / 12, 10, 10000)), -1037.03);
  assert.equal(cents(call('PMT', 0.08 / 12, 10, 10000, 0, 1)), -1030.16);
  assert.equal(cents(call('FV', 0.06 / 12, 10, -200, -500, 1)), 2581.4);
  assert.equal(cents(call('FV', 0.12 / 12, 12, -1000)), 12682.5);
  assert.equal(cents(call('FV', 0.11 / 12, 35, -2000, 0, 1)), 82846.25);
  assert.equal(cents(call('PV', 0.08 / 12, 240, 500, 0, 0)), -59777.15);
  assert.equal(cents(call('NPER', 0.12 / 12, -100, -1000, 10000)), 60.08);
  assert.equal(cents(call('NPER', 0.12 / 12, -100, -1000)), -9.58);
  // The four functions describe one equation: what PMT gives, FV, PV and NPER take back.
  const pmt = call('PMT', 0.004, 60, 20000, -3000, 1) as number;
  check([
    [-3000, 'FV', 0.004, 60, pmt, 20000, 1],
    [20000, 'PV', 0.004, 60, pmt, -3000, 1],
    [60, 'NPER', 0.004, pmt, 20000, -3000, 1],
  ]);
  const rate = call('RATE', 60, pmt, 20000, -3000, 1) as number;
  assert.ok(Math.abs(rate - 0.004) < 1e-9);
});

test('RATE and IRR iterate to the root or give #NUM!', () => {
  const rate = call('RATE', 48, -200, 8000) as number;
  assert.equal(Math.round(rate * 1e6) / 1e6, 0.007701);
  assert.ok(Math.abs((call('RATE', 10, -1000, 10000) as number) - 0) < 1e-9);
  assert.ok(Math.abs((call('RATE', 1, 0, -1000, 1100) as number) - 0.1) < 1e-9);
  assert.ok(Math.abs((call('RATE', 12, -100, 1000, 0, 0, 0.5) as number) - 0.029228540769138) < 1e-9);
  check([
    // No rate makes payments of the same sign balance.
    ['#NUM!', 'RATE', 10, 100, 1000],
    ['#NUM!', 'RATE', 0, -100, 1000],
    ['#NUM!', 'RATE', 10, -100, 1000, 0, 0, -2],
    ['#VALUE!', 'RATE', 'x', -100, 1000],
  ]);
  const flows = arr([[-70000, 12000, 15000, 18000, 21000, 26000]]);
  const round4 = (value: Value): number => Math.round((value as number) * 1e4) / 1e4;
  assert.equal(round4(call('IRR', flows)), 0.0866);
  assert.equal(round4(call('IRR', arr([[-70000, 12000, 15000, 18000, 21000]]))), -0.0212);
  assert.equal(round4(call('IRR', arr([[-70000, 12000, 15000]]), -0.1)), -0.4435);
  assert.ok(Math.abs((call('IRR', arr([[-100, 110]])) as number) - 0.1) < 1e-9);
  // A poor guess still finds the root.
  assert.equal(round4(call('IRR', flows, 5)), 0.0866);
  assert.equal(round4(call('IRR', flows, -0.9)), 0.0866);
  context = makeContext({ A1: -100, A2: 'note', A3: 60, A4: 60 });
  const fromRange = call('IRR', ref('A1:A4')) as number;
  assert.ok(Math.abs(-100 + 60 / (1 + fromRange) + 60 / Math.pow(1 + fromRange, 2)) < 1e-6);
  check([
    ['#NUM!', 'IRR', arr([[100, 200]])],
    ['#NUM!', 'IRR', arr([[-100, -200]])],
    ['#NUM!', 'IRR', arr([['a', 'b']])],
    ['#NUM!', 'IRR', arr([[-100, 110]]), -1],
    ['#N/A', 'IRR', arr([[-100, err('#N/A')]])],
  ]);
});

test('NPV, IPMT, PPMT and SLN', () => {
  context = makeContext({ A1: -10000, A2: 3000, A3: 'skip', A4: 4200, A5: 6800 });
  check([
    [1188.4434123352207, 'NPV', 0.1, -10000, 3000, 4200, 6800],
    [1188.4434123352207, 'NPV', 0.1, ref('A1:A5')],
    [1188.4434123352207, 'NPV', 0.1, arr([[-10000, 3000]]), 4200, 6800],
    [100, 'NPV', 0, 40, 60],
    ['#DIV/0!', 'NPV', -1, 100],
    ['#VALUE!', 'NPV', 'x', 100],
    [-66.66666666666667, 'IPMT', 0.1 / 12, 1, 36, 8000],
    [-75.62318600836664, 'PPMT', 0.1 / 12, 1, 24, 2000],
    [0, 'IPMT', 0.1, 1, 3, 8000, 0, 1],
    ['#NUM!', 'IPMT', 0.1, 0, 3, 8000],
    ['#NUM!', 'IPMT', 0.1, 4, 3, 8000],
    ['#NUM!', 'PPMT', 0.1, 4, 3, 8000],
    [2250, 'SLN', 30000, 7500, 10],
    ['#DIV/0!', 'SLN', 30000, 7500, 0],
  ]);
  const cents = (value: Value): number => Math.round((value as number) * 100) / 100;
  assert.equal(cents(call('IPMT', 0.1, 3, 3, 8000)), -292.45);
  assert.equal(cents(call('PPMT', 0.08, 10, 10, 200000)), -27598.05);
  // Interest and principal add up to the payment in every period, and the principal parts to the loan.
  let principal = 0;
  for (let per = 1; per <= 24; per++) {
    const interest = call('IPMT', 0.01, per, 24, 5000) as number;
    const part = call('PPMT', 0.01, per, 24, 5000) as number;
    assert.ok(Math.abs(interest + part - (call('PMT', 0.01, 24, 5000) as number)) < 1e-9);
    principal += part;
  }
  assert.ok(Math.abs(principal + 5000) < 1e-8);
  for (let per = 1; per <= 12; per++) {
    const interest = call('IPMT', 0.02, per, 12, 9000, 500, 1) as number;
    const part = call('PPMT', 0.02, per, 12, 9000, 500, 1) as number;
    assert.ok(Math.abs(interest + part - (call('PMT', 0.02, 12, 9000, 500, 1) as number)) < 1e-9);
  }
});
