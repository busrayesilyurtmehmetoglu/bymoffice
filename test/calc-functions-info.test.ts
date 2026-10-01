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

test('kind tests', () => {
  context = makeContext({ A1: 1, A2: 'x', A3: true, A5: err('#N/A'), A6: '' });
  const kinds: Value[] = [1, 'x', true, null, err('#N/A'), err('#DIV/0!'), ''];
  const table: [string, boolean[]][] = [
    ['ISNUMBER', [true, false, false, false, false, false, false]],
    ['ISTEXT', [false, true, false, false, false, false, true]],
    ['ISNONTEXT', [true, false, true, true, true, true, false]],
    ['ISLOGICAL', [false, false, true, false, false, false, false]],
    ['ISERROR', [false, false, false, false, true, true, false]],
    ['ISERR', [false, false, false, false, false, true, false]],
    ['ISNA', [false, false, false, false, true, false, false]],
  ];
  for (const [name, expected] of table) {
    for (let i = 0; i < kinds.length; i++) assert.equal(call(name, kinds[i]), expected[i], `${name} #${i}`);
  }
  check([
    [true, 'ISNUMBER', ref('A1')],
    [false, 'ISNUMBER', ref('A4')],
    [true, 'ISNONTEXT', ref('A4')],
    [true, 'ISNA', ref('A5')],
    // A range gives one answer per cell.
    [[[true], [false], [false], [false], [false], [false]], 'ISNUMBER', ref('A1:A6')],
    [[[false], [true], [false], [false], [false], [true]], 'ISTEXT', ref('A1:A6')],
    [[[false, true]], 'ISERROR', arr([[1, err('#REF!')]])],
    // Numeric text is still text.
    [false, 'ISNUMBER', '12'],
  ]);
});

test('ISEVEN and ISODD', () => {
  check([
    [true, 'ISEVEN', 2],
    [false, 'ISEVEN', 3],
    [true, 'ISEVEN', 0],
    [true, 'ISEVEN', 2.9],
    [true, 'ISEVEN', -2.5],
    [false, 'ISEVEN', -1],
    [true, 'ISEVEN', '4'],
    [true, 'ISEVEN', null],
    ['#VALUE!', 'ISEVEN', 'x'],
    ['#VALUE!', 'ISEVEN', true],
    ['#N/A', 'ISEVEN', err('#N/A')],
    [true, 'ISODD', 3],
    [false, 'ISODD', 2],
    [true, 'ISODD', -1.9],
    [true, 'ISODD', 5.5],
    ['#VALUE!', 'ISODD', 'x'],
    [[[true, false]], 'ISODD', arr([[1, 2]])],
  ]);
});

test('NA, TRUE, FALSE, ERROR.TYPE and TYPE', () => {
  context = makeContext({ A1: 1, A2: 'x' });
  check([
    ['#N/A', 'NA'],
    [true, 'TRUE'],
    [false, 'FALSE'],
    [1, 'ERROR.TYPE', err('#NULL!')],
    [2, 'ERROR.TYPE', err('#DIV/0!')],
    [3, 'ERROR.TYPE', err('#VALUE!')],
    [4, 'ERROR.TYPE', err('#REF!')],
    [5, 'ERROR.TYPE', err('#NAME?')],
    [6, 'ERROR.TYPE', err('#NUM!')],
    [7, 'ERROR.TYPE', err('#N/A')],
    [9, 'ERROR.TYPE', err('#SPILL!')],
    [14, 'ERROR.TYPE', err('#CALC!')],
    ['#N/A', 'ERROR.TYPE', 1],
    ['#N/A', 'ERROR.TYPE', 'x'],
    [1, 'TYPE', 5],
    [2, 'TYPE', 'x'],
    [4, 'TYPE', true],
    [16, 'TYPE', err('#N/A')],
    [64, 'TYPE', arr([[1, 2]])],
    [64, 'TYPE', ref('A1:A2')],
    [1, 'TYPE', ref('A1')],
    [2, 'TYPE', ref('A2')],
    // An empty cell is a number (zero).
    [1, 'TYPE', ref('A3')],
  ]);
});
