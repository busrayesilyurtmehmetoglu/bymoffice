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

const DATA = [3, 7, 7, 19, 2, 11, 5.5, -4];
const SHEET: Cells = {
  A1: 1, A2: 2, A3: 'x', A4: true, A5: 4, A6: '7',
  B1: 10, B2: 20, B3: 30, B4: 40, B5: 50,
  C1: 'apple', C2: 'Banana', C3: 'cherry', C4: 'apple', C5: '',
  D1: 5, D2: err('#DIV/0!'), D3: 7,
  E1: 3, E2: 7, E3: 7, E4: 19, E5: 2, E6: 11, E7: 5.5, E8: -4,
  F1: 45292, F2: 45366, F3: 45657, F4: false, F5: 0, F6: '0',
};

test('averages, counts and extremes', () => {
  context = makeContext(SHEET);
  check([
    [6.3125, 'AVERAGE', ref('E1:E8')],
    [6.3125, 'AVERAGE', arr([DATA])],
    [7 / 3, 'AVERAGE', ref('A1:A6')],
    [2, 'AVERAGE', 1, '2', true, 4],
    ['#DIV/0!', 'AVERAGE', ref('C1:C5')],
    ['#DIV/0!', 'AVERAGE', ref('D1:D3')],
    ['#VALUE!', 'AVERAGE', 'abc'],
    // AVERAGEA: text in a range is 0, TRUE is 1, empty cells are skipped.
    [8 / 6, 'AVERAGEA', ref('A1:A6')],
    [8 / 6, 'AVERAGEA', ref('A1:A9')],
    [3, 'COUNT', ref('A1:A6')],
    [5, 'COUNT', ref('A1:A6'), '1', true, 'x'],
    [2, 'COUNT', ref('D1:D3')],
    [1, 'COUNT', err('#N/A'), 5],
    [6, 'COUNTA', ref('A1:A9')],
    [5, 'COUNTA', ref('C1:C9')],
    [3, 'COUNTA', ref('D1:D3')],
    [3, 'COUNTA', 1, '', err('#N/A')],
    [4, 'COUNTBLANK', ref('A7:A10')],
    // Empty text counts as blank.
    [5, 'COUNTBLANK', ref('C1:C9')],
    [1048576 - 6, 'COUNTBLANK', ref('A:A')],
    [-4, 'MIN', ref('E1:E8')],
    [19, 'MAX', ref('E1:E8')],
    [1, 'MIN', ref('A1:A6')],
    [4, 'MAX', ref('A1:A6')],
    [0, 'MIN', ref('C1:C5')],
    [0, 'MAX', ref('C1:C5')],
    [-1, 'MIN', 5, '-1', true],
    ['#DIV/0!', 'MAX', ref('D1:D3')],
    [0, 'MINA', ref('A1:A6')],
    [1, 'MAXA', arr([[-5, true, 'x']])],
    [0, 'MAXA', arr([[-5, 'x']])],
    [-5, 'MINA', arr([[-5, true, 'x']])],
  ]);
  // UNSURE: an argument left empty, as in COUNT(1,) and AVERAGE(4,). It is taken as a zero here.
  check([
    [2, 'COUNT', 1, null],
    [2, 'AVERAGE', 4, null],
  ]);
});

test('spread: variance and standard deviation', () => {
  context = makeContext(SHEET);
  const e = ref('E1:E8');
  for (const name of ['STDEV', 'STDEV.S']) check([[6.766184301362179, name, e]]);
  for (const name of ['STDEV.P', 'STDEVP']) check([[6.3291858678664195, name, e]]);
  for (const name of ['VAR', 'VAR.S']) check([[45.78125, name, e]]);
  for (const name of ['VAR.P', 'VARP']) check([[40.05859375, name, e]]);
  check([
    ['#DIV/0!', 'STDEV', 5],
    [0, 'STDEV.P', 5],
    ['#DIV/0!', 'VAR.P', ref('C1:C5')],
    ['#DIV/0!', 'VAR', ref('D1:D3')],
    [0, 'VAR', 2, 2, 2],
    [0.5, 'VAR', 1, 2],
    [1 / 3, 'VAR', 1, true, '2'],
  ]);
});

test('order statistics', () => {
  context = makeContext(SHEET);
  const e = ref('E1:E8');
  check([
    [6.25, 'MEDIAN', e],
    [7, 'MEDIAN', 1, 7, 100],
    ['#NUM!', 'MEDIAN', ref('C1:C5')],
    [7, 'MODE', e],
    [7, 'MODE.SNGL', e],
    // Several values tie: the one met first wins.
    [4, 'MODE', 4, 9, 9, 4],
    ['#N/A', 'MODE', 1, 2, 3],
    [19, 'LARGE', e, 1],
    [7, 'LARGE', e, 3],
    [-4, 'LARGE', e, 8],
    ['#NUM!', 'LARGE', e, 9],
    ['#NUM!', 'LARGE', e, 0],
    [[[19, 11]], 'LARGE', e, arr([[1, 2]])],
    [-4, 'SMALL', e, 1],
    [3, 'SMALL', e, 3],
    ['#NUM!', 'SMALL', e, 9],
    ['#NUM!', 'SMALL', ref('C1:C5'), 1],
    ['#DIV/0!', 'SMALL', ref('D1:D3'), 1],
    [1, 'RANK', 19, e],
    [3, 'RANK', 7, e],
    [3, 'RANK.EQ', 7, e],
    [3.5, 'RANK.AVG', 7, e],
    [8, 'RANK', -4, e],
    [1, 'RANK', -4, e, 1],
    [5, 'RANK', 7, e, 1],
    ['#N/A', 'RANK', 8, e],
    [[[1], [8]], 'RANK', arr([[19], [-4]]), e],
    [22, 'PERCENTILE', ref('B1:B5'), 0.3],
    [10, 'PERCENTILE.INC', ref('B1:B5'), 0],
    [50, 'PERCENTILE.INC', ref('B1:B5'), 1],
    ['#NUM!', 'PERCENTILE', ref('B1:B5'), 1.1],
    ['#NUM!', 'PERCENTILE', ref('C1:C5'), 0.5],
    [15, 'PERCENTILE.EXC', ref('B1:B5'), 0.25],
    [30, 'PERCENTILE.EXC', ref('B1:B5'), 0.5],
    ['#NUM!', 'PERCENTILE.EXC', ref('B1:B5'), 0.1],
    ['#NUM!', 'PERCENTILE.EXC', ref('B1:B5'), 0],
    [-4, 'QUARTILE', e, 0],
    [2.75, 'QUARTILE', e, 1],
    [6.25, 'QUARTILE.INC', e, 2],
    [8, 'QUARTILE', e, 3],
    [19, 'QUARTILE', e, 4],
    ['#NUM!', 'QUARTILE', e, 5],
    [2.25, 'QUARTILE.EXC', e, 1],
    [10, 'QUARTILE.EXC', e, 3],
    ['#NUM!', 'QUARTILE.EXC', e, 4],
    [5.476986969656962, 'GEOMEAN', 4, 5, 8, 7, 11, 4, 3],
    ['#NUM!', 'GEOMEAN', 4, 0],
    ['#NUM!', 'GEOMEAN', 4, -1],
  ]);
  // UNSURE: a fractional k in LARGE and SMALL is cut to a whole number here; Excel may round it up.
  check([[11, 'LARGE', e, 2.9]]);
});

test('criteria: operators, wildcards, numbers as text, blanks, booleans, errors, dates', () => {
  context = makeContext(SHEET);
  check([
    [2, 'COUNTIF', ref('C1:C5'), 'apple'],
    [2, 'COUNTIF', ref('C1:C5'), '=APPLE'],
    [3, 'COUNTIF', ref('C1:C5'), '<>apple'],
    [3, 'COUNTIF', ref('C1:C5'), '*a*'],
    [1, 'COUNTIF', ref('C1:C5'), '?anana'],
    [1, 'COUNTIF', ref('C1:C5'), 'ch*y'],
    [0, 'COUNTIF', ref('C1:C5'), 'a?'],
    // "*" takes every text cell, the empty text included.
    [5, 'COUNTIF', ref('C1:C5'), '*'],
    [4, 'COUNTIF', ref('C1:C5'), '?*'],
    [2, 'COUNTIF', ref('C1:C5'), '>b'],
    // UNSURE: the empty text in C5 is taken as text that sorts before "banana".
    [4, 'COUNTIF', ref('C1:C5'), '<=banana'],
    [2, 'COUNTIF', ref('B1:B5'), '>30'],
    [3, 'COUNTIF', ref('B1:B5'), '>=30'],
    [2, 'COUNTIF', ref('B1:B5'), '<30'],
    [1, 'COUNTIF', ref('B1:B5'), 30],
    [1, 'COUNTIF', ref('B1:B5'), '30'],
    [1, 'COUNTIF', ref('B1:B5'), '=30'],
    [4, 'COUNTIF', ref('B1:B5'), '<>30'],
    [1, 'COUNTIF', ref('B1:B5'), '3E1'],
    // A number meets numeric text and the other way round.
    [1, 'COUNTIF', ref('A1:A6'), 7],
    [1, 'COUNTIF', ref('A1:A6'), '7'],
    // Relations compare numbers only: the text "7" and TRUE are not greater than 1.
    [2, 'COUNTIF', ref('A1:A6'), '>1'],
    [1, 'COUNTIF', ref('A1:A6'), true],
    [1, 'COUNTIF', ref('A1:A6'), 'TRUE'],
    [5, 'COUNTIF', ref('A1:A6'), '<>TRUE'],
    [1, 'COUNTIF', ref('D1:D3'), '#DIV/0!'],
    [0, 'COUNTIF', ref('D1:D3'), '#N/A'],
    [2, 'COUNTIF', ref('D1:D3'), '<>#DIV/0!'],
    // Empty criteria: "" takes empty cells and empty text, "=" only empty cells, "<>" everything that holds something.
    [5, 'COUNTIF', ref('C1:C9'), ''],
    [4, 'COUNTIF', ref('C1:C9'), '='],
    [5, 'COUNTIF', ref('C1:C9'), '<>'],
    // "<>x" also takes empty cells.
    [7, 'COUNTIF', ref('C1:C9'), '<>apple'],
    [1048576 - 2, 'COUNTIF', ref('C:C'), '<>apple'],
    [2, 'COUNTIF', ref('F1:F3'), '>=3/15/2024'],
    [1, 'COUNTIF', ref('F1:F3'), '3/15/2024'],
    [1, 'COUNTIF', ref('F1:F3'), '<2024-03-15'],
    [1, 'COUNTIF', ref('F4:F6'), false],
    [2, 'COUNTIF', ref('F4:F6'), 0],
    [[[2, 1]], 'COUNTIF', ref('C1:C5'), arr([['apple', 'cherry']])],
    ['#VALUE!', 'COUNTIFS', ref('C1:C5'), 'apple', ref('B1:B4'), '>0'],
    [1, 'COUNTIFS', ref('C1:C5'), 'apple', ref('B1:B5'), '>10'],
    [2, 'COUNTIFS', ref('C1:C5'), '<>apple', ref('A1:A5'), '<>2'],
    [1048576 - 3, 'COUNTIFS', ref('C:C'), '<>apple', ref('A:A'), '<>2'],
    ['#VALUE!', 'COUNTIFS', ref('C1:C5'), 'apple', ref('B1:B5')],
  ]);
  context = makeContext({ A1: 'a*c', A2: 'abc', A3: 'a?c', A4: 'a~c', A5: 'AxC' });
  check([
    [5, 'COUNTIF', ref('A1:A5'), 'a*c'],
    [1, 'COUNTIF', ref('A1:A5'), 'a~*c'],
    [1, 'COUNTIF', ref('A1:A5'), 'a~?c'],
    [1, 'COUNTIF', ref('A1:A5'), 'a~~c'],
    [5, 'COUNTIF', ref('A1:A5'), 'a?c'],
  ]);
  // UNSURE: a criterion that is an empty cell. It is read as 0 here.
  context = makeContext({ A1: 0, A2: 5 });
  check([[1, 'COUNTIF', ref('A1:A3'), ref('B9')]]);
});

test('a hostile wildcard pattern is matched in bounded time', () => {
  const text = 'a'.repeat(30000) + 'b';
  context = makeContext({ A1: text });
  const started = Date.now();
  check([
    [0, 'COUNTIF', ref('A1'), '*a*a*a*a*a*a*a*a*a*a*a*a*a*a*a*c'],
    [1, 'COUNTIF', ref('A1'), '*a*a*a*a*a*a*a*a*a*a*a*a*a*a*a*b'],
  ]);
  assert.ok(Date.now() - started < 20000);
});

test('AVERAGEIF, AVERAGEIFS, MINIFS, MAXIFS', () => {
  context = makeContext(SHEET);
  check([
    [35, 'AVERAGEIF', ref('A1:A5'), '>1', ref('B1:B5')],
    [3, 'AVERAGEIF', ref('A1:A5'), '>1'],
    [25, 'AVERAGEIF', ref('C1:C5'), 'apple', ref('B1')],
    ['#DIV/0!', 'AVERAGEIF', ref('A1:A5'), '>100', ref('B1:B5')],
    [35, 'AVERAGEIFS', ref('B1:B5'), ref('A1:A5'), '>=2', ref('C1:C5'), '<>cherry'],
    ['#DIV/0!', 'AVERAGEIFS', ref('B1:B5'), ref('C1:C5'), 'pear'],
    ['#VALUE!', 'AVERAGEIFS', ref('B1:B5'), ref('C1:C4'), 'apple'],
    [10, 'MINIFS', ref('B1:B5'), ref('C1:C5'), 'apple'],
    [40, 'MAXIFS', ref('B1:B5'), ref('C1:C5'), 'apple'],
    [0, 'MAXIFS', ref('B1:B5'), ref('C1:C5'), 'pear'],
    [0, 'MINIFS', ref('B1:B5'), ref('C1:C5'), 'pear'],
    [20, 'MINIFS', ref('B1:B5'), ref('C1:C5'), '<>apple', ref('A1:A5'), '<>x'],
  ]);
});

test('regression over two lists', () => {
  context = makeContext({
    A1: 1, A2: 2, A3: 3, A4: 4, A5: 5, A6: 6,
    B1: 2, B2: 4.1, B3: 5.9, B4: 8.2, B5: 9.8, B6: 12.5,
    C1: 1, C2: 'skip', C3: 3, D1: 5, D2: 100, D3: 9,
  });
  check([
    [0.9982661552420476, 'CORREL', ref('A1:A6'), ref('B1:B6')],
    [2.0542857142857143, 'SLOPE', ref('B1:B6'), ref('A1:A6')],
    [-0.10666666666666647, 'INTERCEPT', ref('B1:B6'), ref('A1:A6')],
    [1, 'CORREL', arr([[1, 2, 3]]), arr([[2, 4, 6]])],
    [-1, 'CORREL', arr([[1, 2, 3]]), arr([[6, 4, 2]])],
    // A pair is used only when both sides are numbers.
    [2, 'SLOPE', ref('D1:D3'), ref('C1:C3')],
    [3, 'INTERCEPT', ref('D1:D3'), ref('C1:C3')],
    ['#N/A', 'CORREL', ref('A1:A6'), ref('B1:B5')],
    ['#DIV/0!', 'CORREL', arr([[1, 1, 1]]), arr([[2, 4, 6]])],
    ['#DIV/0!', 'SLOPE', arr([[2, 4, 6]]), arr([[1, 1, 1]])],
    ['#DIV/0!', 'SLOPE', arr([['a']]), arr([['b']])],
    ['#N/A', 'SLOPE', arr([[1, err('#N/A')]]), arr([[1, 2]])],
  ]);
});

test('order statistics of a million numbers', () => {
  const data: Scalar[] = new Array(1000000);
  for (let i = 0; i < data.length; i++) data[i] = (i * 7919) % 1000003;
  const big: Value = { rows: 1000000, cols: 1, data };
  assert.equal(call('MAX', big), 1000002);
  assert.equal(call('MIN', big), 0);
  assert.equal(call('COUNT', big), 1000000);
  assert.equal(typeof call('MEDIAN', big), 'number');
  assert.equal(call('LARGE', big, 1), 1000002);
  assert.equal(typeof call('STDEV', big), 'number');
  assert.equal(call('COUNTIF', big, '>=0'), 1000000);
});
