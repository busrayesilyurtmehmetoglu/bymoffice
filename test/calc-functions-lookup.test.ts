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

const TABLE: Cells = {
  A1: 10, B1: 'ten', C1: 'x1',
  A2: 20, B2: 'twenty', C2: 'x2',
  A3: 30, B3: 'thirty', C3: 'x3',
  A4: 40, B4: 'forty', C4: 'x4',
  E1: 'pear', F1: 1, E2: 'Apple', F2: 2, E3: 'apricot', F3: 3, E4: 'fig?', F4: 4, E5: 'apple', F5: 5,
  H1: 1, I1: 2, J1: 3, K1: 4,
  H2: 'a', I2: 'b', J2: 'c', K2: 'd',
  M1: 40, M2: 30, M3: 20, M4: 10,
};

test('VLOOKUP and HLOOKUP', () => {
  context = makeContext(TABLE);
  const t = ref('A1:C4');
  check([
    ['twenty', 'VLOOKUP', 20, t, 2, false],
    ['x3', 'VLOOKUP', 30, t, 3, false],
    ['#N/A', 'VLOOKUP', 25, t, 2, false],
    ['#N/A', 'VLOOKUP', '20', t, 2, false],
    // Approximate: the last row not greater than the lookup value.
    ['twenty', 'VLOOKUP', 25, t, 2],
    ['twenty', 'VLOOKUP', 25, t, 2, true],
    ['forty', 'VLOOKUP', 1000, t, 2],
    ['ten', 'VLOOKUP', 10, t, 2],
    ['#N/A', 'VLOOKUP', 5, t, 2],
    ['#REF!', 'VLOOKUP', 20, t, 4, false],
    ['#VALUE!', 'VLOOKUP', 20, t, 0, false],
    ['#N/A', 'VLOOKUP', err('#N/A'), t, 2, false],
    [10, 'VLOOKUP', 10, t, 1, 0],
    // Text is matched without regard to case; an exact search takes wildcards.
    [2, 'VLOOKUP', 'apple', ref('E1:F5'), 2, false],
    [2, 'VLOOKUP', 'ap*', ref('E1:F5'), 2, false],
    [3, 'VLOOKUP', 'apr?cot', ref('E1:F5'), 2, false],
    [4, 'VLOOKUP', 'fig~?', ref('E1:F5'), 2, false],
    ['#N/A', 'VLOOKUP', 'fig~*', ref('E1:F5'), 2, false],
    [1, 'VLOOKUP', '*', ref('E1:F5'), 2, false],
    // A whole-column table costs only its used rows.
    ['thirty', 'VLOOKUP', 30, ref('A:C'), 2, false],
    ['forty', 'VLOOKUP', 99, ref('A:C'), 2, true],
    [[['ten'], ['thirty']], 'VLOOKUP', arr([[10], [30]]), t, 2, false],
    ['twenty', 'VLOOKUP', 20, arr([[10, 'ten'], [20, 'twenty']]), 2, false],
    ['c', 'HLOOKUP', 3, ref('H1:K2'), 2, false],
    ['b', 'HLOOKUP', 2.5, ref('H1:K2'), 2],
    ['#N/A', 'HLOOKUP', 0, ref('H1:K2'), 2],
    ['#REF!', 'HLOOKUP', 3, ref('H1:K2'), 3, false],
    [4, 'HLOOKUP', 'D', arr([['a', 'd'], [1, 4]]), 2, false],
  ]);
  // An empty result cell comes back empty (the caller shows it as 0).
  context = makeContext({ A1: 1, A2: 2, B2: 'x' });
  assert.equal(call('VLOOKUP', 1, ref('A1:B2'), 2, false), null);
  // Empty cells and text inside a sorted numeric column are stepped over.
  context = makeContext({ A1: 1, A2: 'n/a', A4: 5, A5: 9, A8: 'end' });
  check([
    [5, 'VLOOKUP', 6, ref('A1:A8'), 1],
    [9, 'VLOOKUP', 60, ref('A1:A8'), 1],
    [1, 'VLOOKUP', 4, ref('A1:A8'), 1],
    ['end', 'VLOOKUP', 'zzz', ref('A1:A8'), 1],
  ]);
});

test('MATCH and LOOKUP', () => {
  context = makeContext(TABLE);
  check([
    [2, 'MATCH', 20, ref('A1:A4'), 0],
    ['#N/A', 'MATCH', 25, ref('A1:A4'), 0],
    [2, 'MATCH', 25, ref('A1:A4')],
    [2, 'MATCH', 25, ref('A1:A4'), 1],
    [4, 'MATCH', 99, ref('A1:A4'), 1],
    ['#N/A', 'MATCH', 5, ref('A1:A4'), 1],
    // -1: descending data, the smallest value not below the lookup value.
    [2, 'MATCH', 25, ref('M1:M4'), -1],
    [2, 'MATCH', 30, ref('M1:M4'), -1],
    [4, 'MATCH', 5, ref('M1:M4'), -1],
    ['#N/A', 'MATCH', 50, ref('M1:M4'), -1],
    [3, 'MATCH', 3, ref('H1:K1'), 0],
    [2, 'MATCH', 'APPLE', ref('E1:E5'), 0],
    [3, 'MATCH', 'apr*', ref('E1:E5'), 0],
    [4, 'MATCH', 'fig~?', ref('E1:E5'), 0],
    ['#N/A', 'MATCH', 20, ref('A1:C4'), 0],
    [2, 'MATCH', 'b', arr([['a', 'b', 'c']]), 0],
    [[[1, 3]], 'MATCH', arr([[10, 30]]), ref('A1:A4'), 0],
    [3, 'MATCH', 30, ref('A:A'), 0],
    ['thirty', 'LOOKUP', 35, ref('A1:A4'), ref('B1:B4')],
    ['forty', 'LOOKUP', 99, ref('A1:A4'), ref('B1:B4')],
    ['#N/A', 'LOOKUP', 5, ref('A1:A4'), ref('B1:B4')],
    // The result vector may lie the other way.
    ['c', 'LOOKUP', 30, ref('A1:A4'), ref('H2:K2')],
    ['#N/A', 'LOOKUP', 40, ref('A1:A4'), ref('H2:J2')],
    // Array form: first column in, last column out (first row for a wide table).
    ['x2', 'LOOKUP', 25, ref('A1:C4')],
    ['c', 'LOOKUP', 3.5, ref('H1:K2')],
    [30, 'LOOKUP', 35, ref('A1:A4')],
  ]);
  // UNSURE: an empty lookup value finds nothing (#N/A) here.
  check([['#N/A', 'MATCH', null, ref('A1:A4'), 0]]);
});

test('XLOOKUP and XMATCH', () => {
  context = makeContext(TABLE);
  const keys = ref('A1:A4');
  check([
    ['twenty', 'XLOOKUP', 20, keys, ref('B1:B4')],
    ['#N/A', 'XLOOKUP', 25, keys, ref('B1:B4')],
    ['none', 'XLOOKUP', 25, keys, ref('B1:B4'), 'none'],
    ['twenty', 'XLOOKUP', 25, keys, ref('B1:B4'), 'none', -1],
    ['thirty', 'XLOOKUP', 25, keys, ref('B1:B4'), 'none', 1],
    ['none', 'XLOOKUP', 5, keys, ref('B1:B4'), 'none', -1],
    ['ten', 'XLOOKUP', 5, keys, ref('B1:B4'), 'none', 1],
    // Next smaller and next larger also work in unsorted data.
    [3, 'XLOOKUP', 'b', ref('E1:E5'), ref('F1:F5'), 0, -1],
    [4, 'XLOOKUP', 'b', ref('E1:E5'), ref('F1:F5'), 0, 1],
    // Wildcards need match mode 2.
    ['#N/A', 'XLOOKUP', 'apr*', ref('E1:E5'), ref('F1:F5')],
    [3, 'XLOOKUP', 'apr*', ref('E1:E5'), ref('F1:F5'), 0, 2],
    // Search from the last.
    [5, 'XLOOKUP', 'apple', ref('E1:E5'), ref('F1:F5'), 0, 0, -1],
    [2, 'XLOOKUP', 'apple', ref('E1:E5'), ref('F1:F5'), 0, 0, 1],
    // Binary search, ascending and descending.
    ['thirty', 'XLOOKUP', 30, keys, ref('B1:B4'), 'none', 0, 2],
    ['none', 'XLOOKUP', 25, keys, ref('B1:B4'), 'none', 0, 2],
    ['twenty', 'XLOOKUP', 25, keys, ref('B1:B4'), 'none', -1, 2],
    ['thirty', 'XLOOKUP', 25, keys, ref('B1:B4'), 'none', 1, 2],
    [3, 'XLOOKUP', 20, ref('M1:M4'), arr([[1], [2], [3], [4]]), 0, 0, -2],
    [3, 'XLOOKUP', 25, ref('M1:M4'), arr([[1], [2], [3], [4]]), 0, -1, -2],
    [2, 'XLOOKUP', 25, ref('M1:M4'), arr([[1], [2], [3], [4]]), 0, 1, -2],
    // A return range of several columns gives the whole row.
    [[['twenty', 'x2']], 'XLOOKUP', 20, keys, ref('B1:C4')],
    ['c', 'XLOOKUP', 3, ref('H1:K1'), ref('H2:K2')],
    [[['x3'], ['c']], 'XLOOKUP', 3, ref('H1:K1'), arr([[1, 2, 'x3', 4], ['a', 'b', 'c', 'd']])],
    [[['ten', 'thirty', 'none']], 'XLOOKUP', arr([[10, 30, 31]]), keys, ref('B1:B4'), 'none'],
    ['#VALUE!', 'XLOOKUP', 20, keys, ref('B1:B3')],
    ['#VALUE!', 'XLOOKUP', 20, ref('A1:B4'), ref('B1:B4')],
    ['#VALUE!', 'XLOOKUP', 20, keys, ref('B1:B4'), 'none', 3],
    ['#VALUE!', 'XLOOKUP', 20, keys, ref('B1:B4'), 'none', 0, 0],
    ['#N/A', 'XLOOKUP', err('#N/A'), keys, ref('B1:B4')],
    ['thirty', 'XLOOKUP', 30, ref('A:A'), ref('B:B')],
    [2, 'XMATCH', 20, keys],
    ['#N/A', 'XMATCH', 25, keys],
    [2, 'XMATCH', 25, keys, -1],
    [3, 'XMATCH', 25, keys, 1],
    [3, 'XMATCH', 'apr*', ref('E1:E5'), 2],
    [5, 'XMATCH', 'apple', ref('E1:E5'), 0, -1],
    [2, 'XMATCH', 30, ref('M1:M4'), 0, -2],
    ['#VALUE!', 'XMATCH', 20, ref('A1:B4')],
    ['#VALUE!', 'XMATCH', 20, keys, 5],
  ]);
});

test('TRANSPOSE, UNIQUE, SORT and FILTER return arrays', () => {
  context = makeContext({
    A1: 'b', B1: 2, A2: 'a', B2: 1, A3: 'B', B3: 2, A4: 'c', B4: 3, A5: 'a', B5: 9,
    D1: 3, D2: 'x', D3: 1, D5: true, D6: 2, D7: err('#N/A'),
  });
  check([
    [[['b', 'a', 'B', 'c', 'a'], [2, 1, 2, 3, 9]], 'TRANSPOSE', ref('A1:B5')],
    [[[1], [2]], 'TRANSPOSE', arr([[1, 2]])],
    [[[5]], 'TRANSPOSE', 5],
    [[['b'], ['a'], ['c']], 'UNIQUE', ref('A1:A5')],
    // Rows are compared as a whole.
    [[['b', 2], ['a', 1], ['c', 3], ['a', 9]], 'UNIQUE', ref('A1:B5')],
    [[['c']], 'UNIQUE', ref('A1:A5'), false, true],
    [[[1, 2]], 'UNIQUE', arr([[1, 2, 1]]), true],
    [[[1, '1', true]], 'UNIQUE', arr([[1, '1', true, 1]]), true],
    ['#CALC!', 'UNIQUE', arr([[1], [1]]), false, true],
    [[['a'], ['a'], ['b'], ['B'], ['c']], 'SORT', ref('A1:A5')],
    [[['c'], ['b'], ['B'], ['a'], ['a']], 'SORT', ref('A1:A5'), 1, -1],
    [[['a', 1], ['b', 2], ['B', 2], ['c', 3], ['a', 9]], 'SORT', ref('A1:B5'), 2],
    [[['a', 9], ['c', 3], ['b', 2], ['B', 2], ['a', 1]], 'SORT', ref('A1:B5'), 2, -1],
    [[[1, 2, 3]], 'SORT', arr([[3, 1, 2]]), 1, 1, true],
    ['#VALUE!', 'SORT', ref('A1:B5'), 3],
    ['#VALUE!', 'SORT', ref('A1:B5'), 1, 2],
    [[['b', 2], ['B', 2]], 'FILTER', ref('A1:B5'), arr([[1], [0], [true], [false], [0]])],
    [[['b'], ['a']], 'FILTER', ref('A1:A2'), ref('B1:B2')],
    [[[1, 3]], 'FILTER', arr([[1, 2, 3]]), arr([[true, false, true]])],
    ['#CALC!', 'FILTER', ref('A1:A2'), arr([[0], [0]])],
    ['empty', 'FILTER', ref('A1:A2'), arr([[0], [0]]), 'empty'],
    ['#VALUE!', 'FILTER', ref('A1:A5'), arr([[1], [0]])],
    ['#VALUE!', 'FILTER', ref('A1:A2'), arr([['x'], [1]])],
    ['#N/A', 'FILTER', ref('A1:A2'), arr([[err('#N/A')], [1]])],
  ]);
  // UNSURE: order across kinds in SORT: numbers, then text, then booleans, then errors; empty cells last either way.
  check([
    [[[1], [2], [3], ['x'], [true], ['#N/A'], [null]], 'SORT', ref('D1:D7')],
    [[['#N/A'], [true], ['x'], [3], [2], [1], [null]], 'SORT', ref('D1:D7'), 1, -1],
  ]);
});

test('exact lookups in 100,000 rows stay usable', () => {
  const rows = 100000;
  const data: Scalar[] = new Array(rows * 2);
  for (let i = 0; i < rows; i++) {
    data[i * 2] = 'key' + i;
    data[i * 2 + 1] = i;
  }
  const table: Value = { rows, cols: 2, data };
  assert.equal(call('VLOOKUP', 'key99999', table, 2, false), 99999);
  assert.equal(plain(call('VLOOKUP', 'missing', table, 2, false)), '#N/A');
  const numbers: Scalar[] = new Array(rows);
  for (let i = 0; i < rows; i++) numbers[i] = i * 2;
  const sorted: Value = { rows, cols: 1, data: numbers };
  // Approximate lookups are binary: a thousand of them cost next to nothing.
  const started = Date.now();
  for (let i = 0; i < 1000; i++) assert.equal(call('MATCH', i * 199 + 1, sorted, 1), Math.floor((i * 199 + 1) / 2) + 1);
  assert.ok(Date.now() - started < 1000);
  assert.equal(call('XMATCH', 199998, sorted, 0, 2), rows);
});
