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

const SHEET: Cells = {
  A1: 1, A2: 2, A3: 'x', A4: true, A5: 4, A6: '7',
  B1: 10, B2: 20, B3: 30, B4: 40, B5: 50,
  C1: 'apple', C2: 'Banana', C3: 'cherry', C4: 'apple', C5: '',
  D1: 5, D2: err('#DIV/0!'), D3: 7,
};

test('SUM and its relatives follow the argument rules of aggregates', () => {
  context = makeContext(SHEET);
  check([
    [6, 'SUM', 1, 2, 3],
    // Inside a range text, booleans and numeric text are skipped ...
    [7, 'SUM', ref('A1:A6')],
    // ... given directly they are converted.
    [11, 'SUM', ref('A1:A5'), '3', true],
    [45293, 'SUM', '1/2/2024'],
    ['#VALUE!', 'SUM', 1, 'abc'],
    [3, 'SUM', arr([[1, '5', true, 2]])],
    [1, 'SUM', 1, null],
    ['#DIV/0!', 'SUM', ref('D1:D3')],
    ['#N/A', 'SUM', 1, err('#N/A')],
    [7, 'SUM', ref('A:A')],
    [157, 'SUM', ref('A:B')],
    [0, 'SUM', ref('F1:F9')],
    [30, 'SUMSQ', 1, 2, 3, 4],
    [21, 'SUMSQ', ref('A1:A5')],
    [24, 'PRODUCT', 2, 3, 4],
    [8, 'PRODUCT', ref('A1:A6')],
    [0, 'PRODUCT', ref('F1:F9')],
    ['#NUM!', 'PRODUCT', 1e200, 1e200],
    [250, 'SUMPRODUCT', ref('A1:A5'), ref('B1:B5')],
    [14, 'SUMPRODUCT', arr([[1, 2, 3]]), arr([[1, 2, 3]])],
    [6, 'SUMPRODUCT', arr([[1, 2, 3]])],
    ['#VALUE!', 'SUMPRODUCT', ref('A1:A5'), ref('B1:B4')],
    ['#DIV/0!', 'SUMPRODUCT', ref('D1:D3'), ref('B1:B3')],
    [250, 'SUMPRODUCT', ref('A:A'), ref('B:B')],
  ]);
});

test('SUMIF and SUMIFS', () => {
  context = makeContext(SHEET);
  check([
    [70, 'SUMIF', ref('A1:A5'), '>1', ref('B1:B5')],
    [6, 'SUMIF', ref('A1:A5'), '>1'],
    // The sum range takes the shape of the criteria range from its top left cell.
    [50, 'SUMIF', ref('C1:C5'), 'apple', ref('B1')],
    [50, 'SUMIF', ref('C1:C5'), 'APPLE', ref('B1:B5')],
    [20, 'SUMIF', ref('C1:C5'), '*an*', ref('B1:B5')],
    [100, 'SUMIF', ref('C1:C5'), '<>apple', ref('B1:B5')],
    [20, 'SUMIF', ref('A1:A5'), 2, ref('B1:B5')],
    // A criterion matching empty cells reaches the sum cells beside them.
    [50, 'SUMIF', ref('C1:C5'), '', ref('B1:B5')],
    [0, 'SUMIF', ref('A1:A5'), '>100', ref('B1:B5')],
    [[[10], [20]], 'SUMIF', ref('A1:A5'), arr([[1], [2]]), ref('B1:B5')],
    [70, 'SUMIFS', ref('B1:B5'), ref('A1:A5'), '>=2', ref('C1:C5'), '<>cherry'],
    [10, 'SUMIFS', ref('B1:B5'), ref('C1:C5'), 'apple', ref('A1:A5'), 1],
    ['#VALUE!', 'SUMIFS', ref('B1:B5'), ref('A1:A4'), 1],
    ['#DIV/0!', 'SUMIF', ref('B1:B3'), '>0', ref('D1:D3')],
    [12, 'SUMIF', ref('B1:B3'), '<>20', ref('D1:D3')],
    [150, 'SUMIF', ref('B:B'), '>0'],
  ]);
});

test('rounding is half away from zero and immune to floating point noise', () => {
  check([
    [1.01, 'ROUND', 1.005, 2],
    [3, 'ROUND', 2.5, 0],
    [-3, 'ROUND', -2.5, 0],
    [2.2, 'ROUND', 2.15, 1],
    [2.1, 'ROUND', 2.149, 1],
    [-1.48, 'ROUND', -1.475, 2],
    [20, 'ROUND', 21.5, -1],
    [1000, 'ROUND', 626.3, -3],
    [0, 'ROUND', 1.98, -1],
    [-100, 'ROUND', -50.55, -2],
    [1200, 'ROUND', 1234.5, -2],
    [0, 'ROUND', -0.4, 0],
    [1.5, 'ROUND', 1.5, 400],
    [0, 'ROUND', 123, -400],
    [1e300, 'ROUND', 1e300, 2],
    [4, 'ROUNDUP', 3.2, 0],
    [77, 'ROUNDUP', 76.9, 0],
    [3.142, 'ROUNDUP', 3.14159, 3],
    [-3.2, 'ROUNDUP', -3.14159, 1],
    [31500, 'ROUNDUP', 31415.92654, -2],
    [3.3, 'ROUNDUP', 1.1 * 3, 1],
    [3, 'ROUNDDOWN', 3.2, 0],
    [76, 'ROUNDDOWN', 76.9, 0],
    [3.141, 'ROUNDDOWN', 3.14159, 3],
    [-3.1, 'ROUNDDOWN', -3.14159, 1],
    [31400, 'ROUNDDOWN', 31415.92654, -2],
    [-3, 'ROUNDDOWN', -3.7, 0],
    [8, 'TRUNC', 8.9],
    [-8, 'TRUNC', -8.9],
    [0, 'TRUNC', 0.45],
    [1.23, 'TRUNC', 1.239, 2],
    [8, 'INT', 8.9],
    [-9, 'INT', -8.9],
    [9, 'MROUND', 10, 3],
    [-9, 'MROUND', -10, -3],
    [1.4, 'MROUND', 1.3, 0.2],
    ['#NUM!', 'MROUND', 5, -2],
    [0, 'MROUND', 5, 0],
    [3, 'CEILING', 2.5, 1],
    [-4, 'CEILING', -2.5, -2],
    [-2, 'CEILING', -2.5, 2],
    [1.5, 'CEILING', 1.5, 0.1],
    [0.24, 'CEILING', 0.234, 0.01],
    [0.3, 'CEILING', 0.3, 0.1],
    ['#NUM!', 'CEILING', 2.5, -2],
    [0, 'CEILING', 2.5, 0],
    [2, 'FLOOR', 3.7, 2],
    [-2, 'FLOOR', -2.5, -2],
    [-4, 'FLOOR', -2.5, 2],
    [1.5, 'FLOOR', 1.58, 0.1],
    [0.23, 'FLOOR', 0.234, 0.01],
    ['#NUM!', 'FLOOR', 2.5, -2],
    ['#DIV/0!', 'FLOOR', 2.5, 0],
    [25, 'CEILING.MATH', 24.3, 5],
    [7, 'CEILING.MATH', 6.7],
    [-8, 'CEILING.MATH', -8.1, 2],
    [-6, 'CEILING.MATH', -5.5, 2, -1],
    [20, 'FLOOR.MATH', 24.3, 5],
    [6, 'FLOOR.MATH', 6.7],
    [-10, 'FLOOR.MATH', -8.1, 2],
    [-4, 'FLOOR.MATH', -5.5, 2, -1],
    [2, 'EVEN', 1.5],
    [4, 'EVEN', 3],
    [2, 'EVEN', 2],
    [-2, 'EVEN', -1],
    [0, 'EVEN', 0],
    [3, 'ODD', 1.5],
    [3, 'ODD', 3],
    [3, 'ODD', 2],
    [-1, 'ODD', -1],
    [-3, 'ODD', -2],
    [1, 'ODD', 0],
  ]);
  // UNSURE: FLOOR(0, 0). Excel documents #DIV/0! for a zero significance; this engine answers 0 when the number is 0 too.
  check([[0, 'FLOOR', 0, 0]]);
});

test('arithmetic functions', () => {
  check([
    [2, 'ABS', -2],
    [[[1, 2]], 'ABS', arr([[1, -2]])],
    [2, 'ABS', '-2'],
    [1, 'ABS', true],
    [0, 'ABS', null],
    ['#VALUE!', 'ABS', 'two'],
    ['#N/A', 'ABS', err('#N/A')],
    [1, 'SIGN', 10],
    [0, 'SIGN', 0],
    [-1, 'SIGN', -0.00001],
    // MOD takes the sign of the divisor.
    [1, 'MOD', 3, 2],
    [1, 'MOD', -3, 2],
    [-1, 'MOD', 3, -2],
    [-1, 'MOD', -3, -2],
    [0.5, 'MOD', 5.5, 1],
    ['#DIV/0!', 'MOD', 3, 0],
    [2, 'QUOTIENT', 5, 2],
    [1, 'QUOTIENT', 4.5, 3.1],
    [-3, 'QUOTIENT', -10, 3],
    ['#DIV/0!', 'QUOTIENT', 1, 0],
    [25, 'POWER', 5, 2],
    [2, 'POWER', 4, 0.5],
    [0.25, 'POWER', 2, -2],
    ['#NUM!', 'POWER', 0, 0],
    ['#DIV/0!', 'POWER', 0, -1],
    ['#NUM!', 'POWER', -8, 0.5],
    ['#NUM!', 'POWER', 10, 400],
    [4, 'SQRT', 16],
    ['#NUM!', 'SQRT', -16],
    [7.38905609893065, 'EXP', 2],
    ['#NUM!', 'EXP', 1000],
    [4.454347296253507, 'LN', 86],
    ['#NUM!', 'LN', 0],
    ['#NUM!', 'LN', -1],
    [1, 'LOG', 10],
    [3, 'LOG', 8, 2],
    [4.454347342888286, 'LOG', 86, 2.7182818],
    ['#NUM!', 'LOG', 0],
    ['#NUM!', 'LOG', 8, -2],
    ['#DIV/0!', 'LOG', 8, 1],
    [5, 'LOG10', 100000],
    ['#NUM!', 'LOG10', 0],
    [Math.PI, 'PI'],
  ]);
});

test('trigonometry', () => {
  check([
    [0, 'SIN', 0],
    [1, 'SIN', Math.PI / 2],
    [1, 'COS', 0],
    [1, 'TAN', 0.7853981633974483],
    [-0.5235987755982988, 'ASIN', -0.5],
    ['#NUM!', 'ASIN', 2],
    [2.0943951023931957, 'ACOS', -0.5],
    ['#NUM!', 'ACOS', -1.5],
    [0.7853981633974483, 'ATAN', 1],
    // ATAN2 takes x first, then y.
    [0.7853981633974483, 'ATAN2', 1, 1],
    [Math.PI, 'ATAN2', -1, 0],
    [Math.PI / 2, 'ATAN2', 0, 1],
    ['#DIV/0!', 'ATAN2', 0, 0],
    [1.1752011936438014, 'SINH', 1],
    [1.5430806348152437, 'COSH', 1],
    [0.46211715726000974, 'TANH', 0.5],
    ['#NUM!', 'COSH', 1000],
    [180, 'DEGREES', Math.PI],
    [4.71238898038469, 'RADIANS', 270],
  ]);
});

test('combinatorics and integers', () => {
  check([
    [120, 'FACT', 5],
    [1, 'FACT', 1.9],
    [1, 'FACT', 0],
    ['#NUM!', 'FACT', -1],
    [7.257415615307994e306, 'FACT', 170],
    ['#NUM!', 'FACT', 171],
    [28, 'COMBIN', 8, 2],
    [2598960, 'COMBIN', 52, 5],
    [1, 'COMBIN', 5, 0],
    [1, 'COMBIN', 5, 5],
    ['#NUM!', 'COMBIN', 3, 5],
    ['#NUM!', 'COMBIN', -1, 1],
    [720, 'PERMUT', 10, 3],
    [970200, 'PERMUT', 100, 3],
    ['#NUM!', 'PERMUT', 3, 5],
    [1, 'GCD', 5, 2],
    [12, 'GCD', 24, 36],
    [1, 'GCD', 7, 1],
    [5, 'GCD', 5, 0],
    [1260, 'GCD', 5040, 1260],
    [4, 'GCD', arr([[8, 12]]), 20],
    ['#NUM!', 'GCD', -5, 2],
    [10, 'LCM', 5, 2],
    [72, 'LCM', 24, 36],
    [180, 'LCM', 12, 18, 20],
    [0, 'LCM', 5, 0],
    ['#NUM!', 'LCM', -5, 2],
    ['#NUM!', 'LCM', 4503599627370497, 4503599627370495],
  ]);
  // UNSURE: PERMUT(0, 0). The documentation says "number <= 0 is #NUM!"; this engine answers 1 (the empty arrangement).
  check([[1, 'PERMUT', 0, 0]]);
});

test('RAND and RANDBETWEEN draw from the context', () => {
  context = makeContext({});
  for (let i = 0; i < 200; i++) {
    const r = call('RAND') as number;
    assert.ok(r >= 0 && r < 1);
    const n = call('RANDBETWEEN', 3, 6) as number;
    assert.ok(n >= 3 && n <= 6 && Math.floor(n) === n);
    assert.equal(call('RANDBETWEEN', 4, 4), 4);
    const f = call('RANDBETWEEN', 0.5, 2.5) as number;
    assert.ok(f === 1 || f === 2);
  }
  check([
    ['#NUM!', 'RANDBETWEEN', 5, 1],
    ['#VALUE!', 'RANDBETWEEN', 'a', 1],
  ]);
  // A generator that breaks its contract cannot put NaN into a cell.
  const broken = { ...context, random: () => NaN };
  assert.equal(FUNCTIONS.get('RAND')!.call([], broken), 0);
  assert.equal(FUNCTIONS.get('RANDBETWEEN')!.call([1, 9], broken), 1);
});

test('SUBTOTAL and AGGREGATE', () => {
  context = makeContext(SHEET);
  const b = ref('B1:B5');
  check([
    [30, 'SUBTOTAL', 1, b],
    [5, 'SUBTOTAL', 2, b],
    [3, 'SUBTOTAL', 2, ref('A1:A6')],
    [6, 'SUBTOTAL', 3, ref('A1:A6')],
    [50, 'SUBTOTAL', 4, b],
    [10, 'SUBTOTAL', 5, b],
    [12000000, 'SUBTOTAL', 6, b],
    [15.811388300841896, 'SUBTOTAL', 7, b],
    [14.142135623730951, 'SUBTOTAL', 8, b],
    [150, 'SUBTOTAL', 9, b],
    [250, 'SUBTOTAL', 10, b],
    [200, 'SUBTOTAL', 11, b],
    // 101 to 111 cannot leave hidden rows out here: they answer as 1 to 11.
    [150, 'SUBTOTAL', 109, b],
    [30, 'SUBTOTAL', 101, b],
    [157, 'SUBTOTAL', 9, b, ref('A1:A6')],
    ['#VALUE!', 'SUBTOTAL', 12, b],
    ['#VALUE!', 'SUBTOTAL', 0, b],
    ['#VALUE!', 'SUBTOTAL', 'x', b],
    ['#DIV/0!', 'SUBTOTAL', 9, ref('D1:D3')],
    [150, 'AGGREGATE', 9, 0, b],
    ['#DIV/0!', 'AGGREGATE', 9, 4, ref('D1:D3')],
    [12, 'AGGREGATE', 9, 6, ref('D1:D3')],
    [6, 'AGGREGATE', 1, 6, ref('D1:D3')],
    [7, 'AGGREGATE', 4, 6, ref('D1:D3')],
    [30, 'AGGREGATE', 12, 4, b],
    ['#N/A', 'AGGREGATE', 13, 4, b],
    [40, 'AGGREGATE', 14, 6, b, 2],
    [5, 'AGGREGATE', 15, 6, ref('D1:D3'), 1],
    [22, 'AGGREGATE', 16, 4, b, 0.3],
    [20, 'AGGREGATE', 17, 4, b, 1],
    [15, 'AGGREGATE', 19, 4, b, 1],
    ['#VALUE!', 'AGGREGATE', 14, 6, b],
    ['#VALUE!', 'AGGREGATE', 20, 6, b],
    ['#VALUE!', 'AGGREGATE', 9, 8, b],
  ]);
});

test('a million numbers are summed without blowing the stack', () => {
  const data: Scalar[] = new Array(1000000);
  for (let i = 0; i < data.length; i++) data[i] = i % 10;
  const big: Value = { rows: 1000000, cols: 1, data };
  assert.equal(call('SUM', big), 4500000);
  assert.equal(call('SUMSQ', big), 28500000);
  assert.equal(call('SUBTOTAL', 4, big), 9);
  assert.equal(call('GCD', big), 1);
  assert.equal(call('SUMPRODUCT', big, big), 28500000);
  const abs = call('ABS', big);
  assert.ok(isArray(abs) && abs.rows === 1000000 && abs.data[999999] === 9);
});
