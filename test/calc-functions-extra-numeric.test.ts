import assert from 'node:assert/strict';
import { test } from 'node:test';
import { EXTRA_FUNCTIONS } from '../src/calc/functions/extra';
import { dateSerial } from '../src/calc/functions/helpers';
import { calcError, isArray, isError, type CalcContext, type ErrorCode, type Scalar, type Value } from '../src/calc/values';
import { MAX_COLS, parseRange } from '../src/model/sheet/types';

// ---------------------------------------------------------------------------
// A sheet made of a small map of cells, enough of a CalcContext to call functions directly.

type Cells = { [name: string]: Scalar };

function makeContext(cells: Cells): CalcContext {
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
    date1904: false,
    now: () => 45366.5,
    random: () => 0.5,
  };
}

let context = makeContext({});

function ref(name: string): Value {
  return { ref: true, sheet: 0, range: parseRange(name)! };
}

function arr(rows: Scalar[][]): Value {
  const data: Scalar[] = [];
  for (const row of rows) for (const value of row) data.push(value);
  return { rows: rows.length, cols: rows[0].length, data };
}

function err(code: ErrorCode): Scalar {
  return calcError(code);
}

function call(name: string, ...args: Value[]): Value {
  const spec = EXTRA_FUNCTIONS[name];
  assert.ok(spec, `${name} is exported`);
  assert.ok(args.length >= spec.min && (spec.max === undefined || args.length <= spec.max), `${name} takes ${args.length} arguments`);
  return spec.call(args, context);
}

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

/** Checks rows of [expected, tolerance, function, ...arguments]: values the Excel help prints with few digits. */
function close(rows: [number, number, string, ...Value[]][]): void {
  for (const [expected, tolerance, name, ...args] of rows) {
    const actual = call(name, ...args);
    const label = `${name}(${args.map((arg) => JSON.stringify(plain(arg))).join(', ')})`;
    assert.ok(typeof actual === 'number' && Math.abs(actual - expected) <= tolerance, `${label}: expected ${expected}, got ${JSON.stringify(plain(actual))}`);
  }
}

function date(year: number, month: number, day: number): number {
  return dateSerial(year, month, day, false);
}

// ---------------------------------------------------------------------------
// Cash flows on dates

const FLOWS = arr([[-10000], [2750], [4250], [3250], [2750]]);
const DATES = arr([[date(2008, 1, 1)], [date(2008, 3, 1)], [date(2008, 10, 30)], [date(2009, 2, 15)], [date(2009, 4, 1)]]);

test('XNPV (Excel help example)', () => {
  close([[2086.647602, 5e-7, 'XNPV', 0.09, FLOWS, DATES]]);
  check([
    // At no interest the flows simply add up.
    [3000, 'XNPV', 0, FLOWS, DATES],
    // One year (365 days) at 10%: 110 is worth 100 today.
    [0, 'XNPV', 0.1, arr([[-100, 110]]), arr([[40000, 40365]])],
    // A single flow stands at its own date: nothing is discounted.
    [110, 'XNPV', 0.1, arr([[110]]), arr([[40365]])],
    // Dates may be text and carry a time of day, which is dropped.
    [0, 'XNPV', 0.1, arr([[-100, 110]]), arr([['2008-01-01', '2008-12-31']])],
    [0, 'XNPV', 0.1, arr([[-100, 110]]), arr([[40000.9, 40365.2]])],
    ['#NUM!', 'XNPV', 0.1, arr([[-100, 110]]), arr([[40000]])],
    ['#NUM!', 'XNPV', 0.1, arr([[-100, 110]]), arr([[40000, 39999]])],
    ['#NUM!', 'XNPV', 0.1, arr([[-100, 110]]), arr([[40000, -1]])],
    ['#VALUE!', 'XNPV', 0.1, arr([[-100, 'x']]), arr([[40000, 40365]])],
    ['#VALUE!', 'XNPV', 0.1, arr([[-100, 110]]), arr([[40000, 'x']])],
    ['#VALUE!', 'XNPV', 'x', FLOWS, DATES],
    ['#N/A', 'XNPV', 0.1, arr([[-100, err('#N/A')]]), arr([[40000, 40365]])],
    ['#DIV/0!', 'XNPV', 0.1, arr([[-100, 110]]), arr([[40000, err('#DIV/0!')]])],
    // UNSURE: Excel's XNPV is believed to refuse negative rates; so does this one.
    ['#NUM!', 'XNPV', -0.1, FLOWS, DATES],
    // UNSURE: an empty cell among the flows or dates is refused rather than read as 0.
    ['#VALUE!', 'XNPV', 0.1, arr([[-100, null]]), arr([[40000, 40365]])],
    ['#VALUE!', 'XNPV', 0.1, arr([[-100, 110]]), arr([[40000, null]])],
  ]);
  context = makeContext({ A1: -100, A2: 110, B1: 40000, B2: 40365 });
  check([[0, 'XNPV', 0.1, ref('A1:A2'), ref('B1:B2')]]);
  context = makeContext({});
});

test('XIRR (Excel help example)', () => {
  // Excel stops at an accuracy of 1e-8 and prints 0.373362535; the root itself is 0.3733625335...
  close([
    [0.373362535, 1e-8, 'XIRR', FLOWS, DATES],
    [0.3733625335188, 1e-12, 'XIRR', FLOWS, DATES, 0.1],
    [0.3733625335188, 1e-12, 'XIRR', FLOWS, DATES, 5],
    [0.3733625335188, 1e-12, 'XIRR', FLOWS, DATES, -0.9],
  ]);
  close([[0, 1e-6, 'XNPV', call('XIRR', FLOWS, DATES) as number, FLOWS, DATES]]);
  check([
    [0.1, 'XIRR', arr([[-100, 110]]), arr([[40000, 40365]])],
    [0.1, 'XIRR', arr([[-100, 110]]), arr([[40000, 40365]]), null],
    // Half of the money lost in a year.
    [-0.5, 'XIRR', arr([[-100, 50]]), arr([[40000, 40365]])],
    // Doubled in half a year (rounded down to 182 days): 2^(365/182) - 1.
    [Math.pow(2, 365 / 182) - 1, 'XIRR', arr([[-100, 200]]), arr([[40000, 40182]])],
    ['#NUM!', 'XIRR', arr([[100, 110]]), arr([[40000, 40365]])],
    ['#NUM!', 'XIRR', arr([[-100, -110]]), arr([[40000, 40365]])],
    ['#NUM!', 'XIRR', arr([[-100, 110]]), arr([[40000, 40365]]), -1],
    ['#NUM!', 'XIRR', arr([[-100, 110, 5]]), arr([[40000, 40365]])],
    ['#NUM!', 'XIRR', arr([[-100, 110]]), arr([[40000, 39000]])],
    ['#VALUE!', 'XIRR', arr([[-100, 'x']]), arr([[40000, 40365]])],
    ['#VALUE!', 'XIRR', arr([[-100, 110]]), arr([[40000, 40365]]), 'x'],
  ]);
  // The rate found makes the present value vanish. XNPV takes no negative rate, so only gains are tried.
  const flows = arr([[-1000, 300, 420, 680, -50, 90]]);
  const dates = arr([[41000, 41100, 41250, 41730, 41800, 42500]]);
  const rate = call('XIRR', flows, dates) as number;
  assert.ok(rate > 0 && rate < 1);
  close([[0, 1e-7, 'XNPV', rate, flows, dates]]);
});

// ---------------------------------------------------------------------------
// Rates, loans and depreciation

test('EFFECT and NOMINAL (Excel help examples)', () => {
  close([
    [0.0535427, 5e-8, 'EFFECT', 0.0525, 4],
    [0.05250032, 5e-9, 'NOMINAL', 0.053543, 4],
  ]);
  check([
    [0.21, 'EFFECT', 0.2, 2],
    [0.21, 'EFFECT', 0.2, 2.9],
    [0.2, 'NOMINAL', 0.21, 2],
    [0.1, 'EFFECT', 0.1, 1],
    [0.1, 'NOMINAL', 0.1, 1],
    ['#NUM!', 'EFFECT', 0, 4],
    ['#NUM!', 'EFFECT', -0.1, 4],
    ['#NUM!', 'EFFECT', 0.1, 0.5],
    ['#NUM!', 'NOMINAL', 0, 4],
    ['#NUM!', 'NOMINAL', 0.1, 0],
    ['#VALUE!', 'EFFECT', 'x', 4],
    [[['#NUM!', '#NUM!']], 'EFFECT', arr([[0, -1]]), arr([[1, 2]])],
  ]);
});

test('CUMIPMT and CUMPRINC (Excel help examples)', () => {
  close([
    [-11135.23213, 5e-6, 'CUMIPMT', 0.09 / 12, 360, 125000, 13, 24, 0],
    [-937.5, 5e-9, 'CUMIPMT', 0.09 / 12, 360, 125000, 1, 1, 0],
    [-934.1071234, 5e-8, 'CUMPRINC', 0.09 / 12, 360, 125000, 13, 24, 0],
    [-68.27827118, 5e-9, 'CUMPRINC', 0.09 / 12, 360, 125000, 1, 1, 0],
  ]);
  // 1000 at 10% over two periods: payments of 576.19..., interest 100 then 52.38...
  const pmt = -(1000 * 0.1 * 1.21) / 0.21;
  close([
    [-100, 1e-9, 'CUMIPMT', 0.1, 2, 1000, 1, 1, 0],
    [pmt + 100, 1e-9, 'CUMPRINC', 0.1, 2, 1000, 1, 1, 0],
    [-(1000 + pmt + 100) * 0.1, 1e-9, 'CUMIPMT', 0.1, 2, 1000, 2, 2, 0],
    // Over the whole loan the principal parts give back what was lent.
    [-1000, 1e-9, 'CUMPRINC', 0.1, 2, 1000, 1, 2, 0],
    [2 * pmt + 1000, 1e-9, 'CUMIPMT', 0.1, 2, 1000, 1, 2, 0],
    // Paid at the start of each period: the first payment bears no interest.
    [0, 1e-9, 'CUMIPMT', 0.1, 2, 1000, 1, 1, 1],
    [pmt / 1.1, 1e-9, 'CUMPRINC', 0.1, 2, 1000, 1, 1, 1],
    [-(1000 + pmt / 1.1) * 0.1, 1e-9, 'CUMIPMT', 0.1, 2, 1000, 2, 2, 1],
    [-1000, 1e-9, 'CUMPRINC', 0.1, 2, 1000, 1, 2, 1],
  ]);
  // Period by period: interest on what is still owed, the rest of the payment is principal.
  for (const type of [0, 1]) {
    const rate = 0.004;
    const periods = 48;
    const growth = Math.pow(1 + rate, periods);
    const payment = (-rate * 20000 * growth) / ((1 + rate * type) * (growth - 1));
    let owed = 20000;
    let interest = 0;
    let principal = 0;
    for (let period = 1; period <= 30; period++) {
      const charge = type === 1 ? (period === 1 ? 0 : -owed * rate) : -owed * rate;
      if (type === 1 && period === 1) owed += payment;
      else owed += payment - charge;
      if (period >= 7) {
        interest += charge;
        principal += payment - charge;
      }
    }
    close([
      [interest, 1e-7, 'CUMIPMT', rate, periods, 20000, 7, 30, type],
      [principal, 1e-7, 'CUMPRINC', rate, periods, 20000, 7, 30, type],
    ]);
  }
  check([
    ['#NUM!', 'CUMIPMT', 0, 360, 125000, 1, 1, 0],
    ['#NUM!', 'CUMIPMT', 0.01, 0, 125000, 1, 1, 0],
    ['#NUM!', 'CUMIPMT', 0.01, 360, 0, 1, 1, 0],
    ['#NUM!', 'CUMIPMT', 0.01, 360, 125000, 0, 1, 0],
    ['#NUM!', 'CUMIPMT', 0.01, 360, 125000, 2, 1, 0],
    ['#NUM!', 'CUMIPMT', 0.01, 360, 125000, 1, 1, 2],
    ['#NUM!', 'CUMPRINC', 0.01, 360, 125000, 1, 1, -1],
    ['#NUM!', 'CUMPRINC', -0.01, 360, 125000, 1, 1, 0],
    ['#VALUE!', 'CUMPRINC', 'x', 360, 125000, 1, 1, 0],
  ]);
  // UNSURE: fractional periods are cut to whole ones, and an end past the last period is not refused.
  close([[-937.5, 5e-9, 'CUMIPMT', 0.09 / 12, 360, 125000, 1.9, 1.9, 0]]);
  // A loan of very many periods costs no more time than a short one.
  close([[-1000, 1e-6, 'CUMPRINC', 1e-9, 1e9, 1000, 1, 1e9, 0]]);
  assert.equal(plain(call('CUMIPMT', 1e-6, 1e12, 1000, 1, 1e12, 0)), '#NUM!');
});

test('DB (Excel help examples)', () => {
  close([
    [186083.33, 5e-3, 'DB', 1000000, 100000, 6, 1, 7],
    [259639.42, 5e-3, 'DB', 1000000, 100000, 6, 2, 7],
    [176814.44, 5e-3, 'DB', 1000000, 100000, 6, 3, 7],
    [120410.64, 5e-3, 'DB', 1000000, 100000, 6, 4, 7],
    [81999.64, 5e-3, 'DB', 1000000, 100000, 6, 5, 7],
    [55841.76, 5e-3, 'DB', 1000000, 100000, 6, 6, 7],
    [15845.1, 5e-3, 'DB', 1000000, 100000, 6, 7, 7],
  ]);
  check([
    // The rate 1 - (1000/10000)^(1/4) = 0.4377 is rounded to 0.438.
    [4380, 'DB', 10000, 1000, 4, 1],
    [5620 * 0.438, 'DB', 10000, 1000, 4, 2],
    [5620 * 0.562 * 0.438, 'DB', 10000, 1000, 4, 3],
    [2190, 'DB', 10000, 1000, 4, 1, 6],
    [7810 * 0.562 * 0.562 * 0.562 * 0.438 * 0.5, 'DB', 10000, 1000, 4, 5, 6],
    ['#NUM!', 'DB', 10000, 1000, 4, 6, 6],
    ['#NUM!', 'DB', 10000, 1000, 4, 0],
    ['#NUM!', 'DB', 10000, 1000, 0, 1],
    ['#NUM!', 'DB', 10000, -1, 4, 1],
    ['#NUM!', 'DB', 10000, 1000, 4, 1, 0],
    ['#NUM!', 'DB', 10000, 1000, 4, 1, 13],
    ['#VALUE!', 'DB', 'x', 1000, 4, 1],
    // UNSURE: what Excel answers for a cost of 0, a salvage above the cost, and the period after
    // the life when the first year was a full one; all are refused here.
    ['#NUM!', 'DB', 0, 0, 4, 1],
    ['#NUM!', 'DB', 1000, 2000, 4, 1],
    ['#NUM!', 'DB', 10000, 1000, 4, 5],
    // UNSURE: a fractional period is cut to a whole one.
    [4380, 'DB', 10000, 1000, 4, 1.9],
  ]);
});

test('DDB (Excel help examples)', () => {
  close([
    [1.31506849, 5e-9, 'DDB', 2400, 300, 10 * 365, 1],
    [40, 1e-9, 'DDB', 2400, 300, 10 * 12, 1, 2],
    [480, 1e-9, 'DDB', 2400, 300, 10, 1, 2],
    [306, 1e-9, 'DDB', 2400, 300, 10, 2, 1.5],
    [22.1225472, 5e-8, 'DDB', 2400, 300, 10, 10],
  ]);
  check([
    // 1000 over 4 years at twice the straight rate: 500, 250, then only down to the salvage of 200.
    [500, 'DDB', 1000, 200, 4, 1],
    [250, 'DDB', 1000, 200, 4, 2],
    [50, 'DDB', 1000, 200, 4, 3],
    [0, 'DDB', 1000, 200, 4, 4],
    // A factor that writes everything off at once.
    [800, 'DDB', 1000, 200, 2, 1, 3],
    [0, 'DDB', 1000, 200, 2, 2, 3],
    ['#NUM!', 'DDB', -1, 200, 4, 1],
    ['#NUM!', 'DDB', 1000, -1, 4, 1],
    ['#NUM!', 'DDB', 1000, 200, 0, 1],
    ['#NUM!', 'DDB', 1000, 200, 4, 0],
    ['#NUM!', 'DDB', 1000, 200, 4, 5],
    ['#NUM!', 'DDB', 1000, 200, 4, 1, 0],
  ]);
});

test('SYD (Excel help examples)', () => {
  close([
    [4090.909091, 5e-7, 'SYD', 30000, 7500, 10, 1],
    [409.0909091, 5e-8, 'SYD', 30000, 7500, 10, 10],
  ]);
  check([
    // 600 over 3 years in shares of 3, 2 and 1 sixths.
    [300, 'SYD', 700, 100, 3, 1],
    [200, 'SYD', 700, 100, 3, 2],
    [100, 'SYD', 700, 100, 3, 3],
    ['#NUM!', 'SYD', 700, 100, 3, 4],
    ['#NUM!', 'SYD', 700, 100, 3, 0],
    ['#NUM!', 'SYD', 700, 100, 0, 1],
    ['#NUM!', 'SYD', 700, -1, 3, 1],
  ]);
});

// ---------------------------------------------------------------------------
// The normal distribution

/** P(Z <= z), from published tables of the standard normal distribution (16 digits). */
const CDF: [number, number][] = [
  [0, 0.5],
  [1, 0.8413447460685429],
  [-1, 0.15865525393145707],
  [1.96, 0.9750021048517795],
  [2.5, 0.9937903346742238],
  [-2.5, 0.006209665325776132],
  [-3, 0.0013498980316300946],
  [3, 0.9986501019683699],
  [-6, 9.865876450376946e-10],
  [-10, 7.619853024160527e-24],
  [0.5, 0.6914624612740131],
  [-0.1, 0.46017216272297107],
];

test('standard normal distribution: accuracy', () => {
  for (const [z, p] of CDF) {
    for (const name of ['NORM.S.DIST', 'NORMSDIST']) {
      const actual = (name === 'NORMSDIST' ? call(name, z) : call(name, z, true)) as number;
      assert.ok(Math.abs(actual - p) <= 1e-14, `${name}(${z}): expected ${p}, got ${actual}`);
      // The lower tail keeps its digits too.
      if (p < 0.01) assert.ok(Math.abs(actual - p) <= 1e-12 * p, `${name}(${z}) relative: expected ${p}, got ${actual}`);
    }
    if (p > 0 && p < 1) {
      for (const name of ['NORM.S.INV', 'NORMSINV']) {
        const actual = call(name, p) as number;
        // Near 1 a probability carries fewer digits of z than a double shows.
        const tolerance = p > 0.99 ? 1e-9 : 1e-11;
        assert.ok(Math.abs(actual - z) <= tolerance, `${name}(${p}): expected ${z}, got ${actual}`);
      }
    }
  }
  close([
    // Excel help examples.
    [0.908788726, 5e-10, 'NORM.S.DIST', 1.333333, true],
    [0.164010148, 5e-10, 'NORM.S.DIST', 1.333333, false],
    [0.908789, 5e-7, 'NORMSDIST', 1.333333],
    [1.3333347, 5e-8, 'NORM.S.INV', 0.908789],
    [1.3333347, 5e-8, 'NORMSINV', 0.908789],
    // Well-known quantiles.
    [1.959963984540054, 1e-12, 'NORM.S.INV', 0.975],
    [1.2815515655446004, 1e-12, 'NORM.S.INV', 0.9],
    [-3.090232306167813, 1e-12, 'NORM.S.INV', 0.001],
    [-6.361340902404056, 1e-11, 'NORM.S.INV', 1e-10],
    [0, 0, 'NORM.S.INV', 0.5],
    // The density at 0 is 1/sqrt(2 pi).
    [0.3989422804014327, 1e-15, 'NORM.S.DIST', 0, false],
    [1, 0, 'NORM.S.DIST', 50, true],
    [0, 0, 'NORM.S.DIST', -50, true],
    [0, 0, 'NORM.S.DIST', 50, false],
  ]);
  // The inverse undoes the distribution across the whole range, the far tail included.
  for (const p of [1e-300, 1e-100, 1e-20, 1e-6, 0.02, 0.024, 0.025, 0.3, 0.5, 0.7, 0.975, 0.999999, 1 - 1e-12]) {
    const z = call('NORM.S.INV', p) as number;
    const back = call('NORM.S.DIST', z, true) as number;
    assert.ok(Math.abs(back - p) <= 1e-12 * Math.min(p, 1) + (p > 0.5 ? 1e-15 : 0), `p = ${p}: z = ${z}, back ${back}`);
  }
  check([
    ['#NUM!', 'NORM.S.INV', 0],
    ['#NUM!', 'NORM.S.INV', 1],
    ['#NUM!', 'NORM.S.INV', -0.1],
    ['#NUM!', 'NORM.S.INV', 1.1],
    ['#VALUE!', 'NORM.S.INV', 'x'],
    ['#VALUE!', 'NORM.S.DIST', 'x', true],
    ['#VALUE!', 'NORM.S.DIST', 1, 'x'],
    ['#N/A', 'NORMSDIST', err('#N/A')],
    [[[0.5, 1]], 'NORMSDIST', arr([[0, 50]])],
  ]);
});

test('NORM.DIST and NORM.INV (Excel help examples)', () => {
  for (const [dist, inv] of [
    ['NORM.DIST', 'NORM.INV'],
    ['NORMDIST', 'NORMINV'],
  ]) {
    close([
      [0.9087888, 5e-8, dist, 42, 40, 1.5, true],
      [0.10934, 5e-6, dist, 42, 40, 1.5, false],
      [42.000002, 5e-7, inv, 0.908789, 40, 1.5],
      [0.5, 0, dist, 40, 40, 1.5, true],
      [0.8413447460685429, 1e-14, dist, 13, 10, 3, true],
      [0.8413447460685429, 1e-14, dist, 13, 10, 3, 1],
      // The density is scaled by the deviation.
      [0.3989422804014327 / 3, 1e-15, dist, 10, 10, 3, false],
      [0.3989422804014327 / 3, 1e-15, dist, 10, 10, 3, 0],
      [10, 0, inv, 0.5, 10, 3],
      [13, 1e-11, inv, 0.8413447460685429, 10, 3],
      [10 - 3 * 3.090232306167813, 1e-11, inv, 0.001, 10, 3],
    ]);
    check([
      ['#NUM!', dist, 42, 40, 0, true],
      ['#NUM!', dist, 42, 40, -1, true],
      ['#NUM!', inv, 0, 40, 1.5],
      ['#NUM!', inv, 1, 40, 1.5],
      ['#NUM!', inv, 0.5, 40, 0],
      ['#VALUE!', dist, 'x', 40, 1.5, true],
      ['#VALUE!', inv, 0.5, 'x', 1.5],
      ['#N/A', dist, 42, err('#N/A'), 1.5, true],
    ]);
  }
});

// ---------------------------------------------------------------------------
// FORECAST, TRIMMEAN, FREQUENCY

test('FORECAST and FORECAST.LINEAR (Excel help example)', () => {
  for (const name of ['FORECAST', 'FORECAST.LINEAR']) {
    close([[10.607253, 5e-7, name, 30, arr([[6, 7, 9, 15, 21]]), arr([[20, 28, 31, 38, 40]])]]);
    check([
      // y = 2x + 1
      [21, name, 10, arr([[3, 5, 7]]), arr([[1, 2, 3]])],
      [1, name, 0, arr([[3], [5], [7]]), arr([[1], [2], [3]])],
      [-9, name, -5, arr([[3, 5, 7]]), arr([[1], [2], [3]])],
      // Pairs without two numbers are left out.
      [21, name, 10, arr([[3, 'x', 5, 7, 100]]), arr([[1, 9, 2, 3, null]])],
      [[[3, 5]], name, arr([[1, 2]]), arr([[3, 5, 7]]), arr([[1, 2, 3]])],
      ['#N/A', name, 10, arr([[3, 5, 7]]), arr([[1, 2]])],
      ['#DIV/0!', name, 10, arr([[3, 5, 7]]), arr([[2, 2, 2]])],
      ['#DIV/0!', name, 10, arr([['a']]), arr([['b']])],
      ['#VALUE!', name, 'x', arr([[3, 5, 7]]), arr([[1, 2, 3]])],
      ['#NUM!', name, 10, arr([[3, err('#NUM!'), 7]]), arr([[1, 2, 3]])],
    ]);
  }
  context = makeContext({ A1: 3, A2: 5, A3: 7, B1: 1, B2: 2, B3: 3 });
  check([[9, 'FORECAST', 4, ref('A1:A3'), ref('B1:B3')]]);
  context = makeContext({});
});

test('TRIMMEAN (Excel help example)', () => {
  close([[3.778, 5e-4, 'TRIMMEAN', arr([[4, 5, 6, 7, 2, 3, 4, 5, 1, 2, 3]]), 0.2]]);
  const ten = arr([[10, 1, 2, 3, 4, 5, 6, 7, 8, 100]]);
  check([
    [34 / 9, 'TRIMMEAN', arr([[4, 5, 6, 7, 2, 3, 4, 5, 1, 2, 3]]), 0.2],
    [14.6, 'TRIMMEAN', ten, 0],
    // 10% of ten points is one point: an odd count is rounded down to none.
    [14.6, 'TRIMMEAN', ten, 0.1],
    [5.625, 'TRIMMEAN', ten, 0.2],
    [5.625, 'TRIMMEAN', ten, 0.39],
    [5.5, 'TRIMMEAN', ten, 0.4],
    // 30% of ten is exactly three points although 10 * 0.3 falls just short in binary.
    [5.625, 'TRIMMEAN', ten, 0.3],
    [5.5, 'TRIMMEAN', ten, 0.9],
    [7, 'TRIMMEAN', 7, 0.5],
    [2, 'TRIMMEAN', arr([[1, 'x', 2, true, 3, null]]), 0],
    ['#NUM!', 'TRIMMEAN', ten, 1],
    ['#NUM!', 'TRIMMEAN', ten, -0.1],
    ['#NUM!', 'TRIMMEAN', arr([['x']]), 0.1],
    ['#VALUE!', 'TRIMMEAN', ten, 'x'],
    ['#DIV/0!', 'TRIMMEAN', arr([[1, err('#DIV/0!')]]), 0.1],
  ]);
  context = makeContext({ A1: 1, A2: 2, A3: 3, A4: 100 });
  check([[2.5, 'TRIMMEAN', ref('A1:A4'), 0.5]]);
  context = makeContext({});
});

test('FREQUENCY (Excel help example)', () => {
  check([
    [[[1], [2], [4], [2]], 'FREQUENCY', arr([[79], [85], [78], [85], [50], [81], [95], [88], [97]]), arr([[70], [79], [89]])],
    // The answer is a column whatever the shape of the bins.
    [[[1], [2], [4], [2]], 'FREQUENCY', arr([[79, 85, 78, 85, 50, 81, 95, 88, 97]]), arr([[70, 79, 89]])],
    // Bins out of order are answered where they stand; a repeated bin gets nothing.
    [[[4], [1], [2], [0], [2]], 'FREQUENCY', arr([[79, 85, 78, 85, 50, 81, 95, 88, 97]]), arr([[89, 70, 79, 79]])],
    // Text, booleans and empty cells are not counted, on either side.
    [[[2], [1]], 'FREQUENCY', arr([[1, 'x', 2, true, null, 9]]), arr([[5, 'y', null]])],
    // No bins: everything is in the rest.
    [[[3]], 'FREQUENCY', arr([[1, 2, 3]]), arr([['none']])],
    [[[0], [0]], 'FREQUENCY', arr([['none']]), arr([[5]])],
    [[[1], [0]], 'FREQUENCY', 5, 5],
    ['#N/A', 'FREQUENCY', arr([[1, err('#N/A')]]), arr([[5]])],
    ['#DIV/0!', 'FREQUENCY', arr([[1, 2]]), arr([[err('#DIV/0!')]])],
  ]);
  context = makeContext({ A1: 1, A2: 2, A3: 3, A4: 4, C1: 2 });
  check([[[[2], [2]], 'FREQUENCY', ref('A1:A6'), ref('C1:C1')]]);
  context = makeContext({});
});
