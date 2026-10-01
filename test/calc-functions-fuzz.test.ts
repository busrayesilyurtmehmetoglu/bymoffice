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

// ---------------------------------------------------------------------------
// Every function with random arguments: nothing throws, every result is a valid value.

function mulberry(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const CODES: ErrorCode[] = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CALC!'];
const NUMBERS = [
  0, 1, -1, 2, 3, 0.5, -0.5, 1.005, 2.5, 7, 10, 12, 100, 255, 256, 365, 1900, 2024, 32767, 32768, 45366, 45366.75,
  2958465, 2958466, -45366, 1e-7, 1e15, 1e16, 9007199254740993, 1e100, -1e100, 1.7976931348623157e308, 5e-324, 0.1,
  -0.1, 4, 5, 6, 11, 17, 21, 60, 61, 101, 109, 14, 19,
];
const TEXTS = [
  '', ' ', 'a', 'abc', 'ABC', 'a*c', '?', '~', '~*', '*', '>5', '<=10', '<>', '=', '<>x', '>=', '12', '-3.5', '1e3',
  '12%', '$1,234.50', '(5)', '1/2/2024', '2024-03-15', '6:30 PM', '25:00', '31/12/2024', 'TRUE', 'false', '#N/A',
  '#DIV/0!', 'Y', 'M', 'D', 'MD', 'YM', 'YD', '0.00', '#,##0', 'yyyy-mm-dd', 'h:mm AM/PM', '[Red]0;(0)', '@', 'General',
  ',', '.', '-', 'İstanbul ığüşöç', '😀', '\ud83d', '\u0000\u0001', 'x'.repeat(300), 'ab'.repeat(20000),
  '9'.repeat(400), '=1+1', "'quoted", '1 2/3', '  padded  ', 'a-b-c', 'line\nbreak',
];
const RANGES = ['A1', 'A1:A9', 'A1:D9', 'B2:C4', 'D1:D9', 'A1:A3', 'B1:B3', 'A9:D12', 'F1:F3', 'A1:D1', 'A:A', 'A:D', 'B:C'];

function fuzzSheet(): Cells {
  const cells: Cells = {};
  const random = mulberry(7);
  const pick: Scalar[] = [1, 2, 3, 2, -4.5, 0, 'a', 'B', 'apple', '12', '', true, false, err('#N/A'), err('#DIV/0!'), 45366, 45000];
  for (let row = 1; row <= 9; row++) {
    for (const col of ['A', 'B', 'C', 'D']) {
      if (random() < 0.8) cells[col + row] = pick[Math.floor(random() * pick.length)];
    }
  }
  cells.F1 = 10;
  cells.F2 = 20;
  cells.F3 = 30;
  return cells;
}

function randomScalar(random: () => number, hostile: boolean): Scalar {
  const kind = random();
  if (kind < 0.4) {
    const r = random();
    if (hostile && r < 0.04) return [NaN, Infinity, -Infinity][Math.floor(random() * 3)];
    if (r < 0.6) return NUMBERS[Math.floor(random() * NUMBERS.length)];
    return (random() - 0.5) * Math.pow(10, Math.floor(random() * 12) - 3);
  }
  if (kind < 0.7) return TEXTS[Math.floor(random() * TEXTS.length)];
  if (kind < 0.8) return random() < 0.5;
  if (kind < 0.9) return null;
  return err(CODES[Math.floor(random() * CODES.length)]);
}

function randomValue(random: () => number, hostile: boolean): Value {
  const kind = random();
  if (kind < 0.55) return randomScalar(random, hostile);
  if (kind < 0.8) {
    const rows = 1 + Math.floor(random() * 4);
    const cols = 1 + Math.floor(random() * 4);
    const data: Scalar[] = [];
    for (let i = 0; i < rows * cols; i++) data.push(randomScalar(random, hostile));
    if (hostile && random() < 0.05) return { rows: 0, cols: 0, data: [] };
    if (hostile && random() < 0.05) return { rows, cols: cols + 1, data };
    return { rows, cols, data };
  }
  if (hostile && random() < 0.05) {
    // A range that is not one: corners the wrong way round, outside the sheet, not whole numbers.
    const broken = [
      { top: 5, left: 0, bottom: 2, right: 0 },
      { top: 0, left: 3, bottom: 4, right: 1 },
      { top: -3, left: -1, bottom: 2, right: 2 },
      { top: 0, left: 0, bottom: 1e12, right: 1e9 },
      { top: 0.5, left: 0, bottom: 2.5, right: 1 },
      { top: NaN, left: 0, bottom: NaN, right: 0 },
    ];
    return { ref: true, sheet: 0, range: broken[Math.floor(random() * broken.length)] };
  }
  return ref(RANGES[Math.floor(random() * RANGES.length)]);
}

function assertScalar(value: unknown, label: string): void {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return;
  if (typeof value === 'number') {
    assert.ok(isFinite(value), `${label}: a number result must be finite, got ${value}`);
    return;
  }
  assert.ok(isError(value) && CODES.indexOf(value.error) >= 0, `${label}: not a value: ${String(value)}`);
}

function assertValue(value: Value, label: string): void {
  if (isArray(value)) {
    assert.ok(value.rows >= 1 && value.cols >= 1, `${label}: an array result has cells`);
    assert.equal(value.data.length, value.rows * value.cols, `${label}: array size`);
    for (let i = 0; i < value.data.length; i++) assertScalar(value.data[i], label);
    return;
  }
  assert.ok(value !== undefined, `${label}: undefined result`);
  assertScalar(value, label);
}

function describe(args: Value[]): string {
  const text = JSON.stringify(args.map(plain), (_key, value) => (typeof value === 'number' && !isFinite(value) ? String(value) : value));
  return text.length > 400 ? text.substring(0, 400) + '...' : text;
}

function fuzz(seed: number, rounds: number, hostile: boolean): void {
  const random = mulberry(seed);
  for (const [name, spec] of FUNCTIONS) {
    for (let round = 0; round < rounds; round++) {
      const most = spec.max === undefined ? spec.min + 3 : spec.max;
      let count = spec.min + Math.floor(random() * (most - spec.min + 1));
      // Now and then the wrong number of arguments: the evaluator checks counts, but a function must survive it.
      if (hostile && random() < 0.05) count = Math.floor(random() * (most + 3));
      const args: Value[] = [];
      for (let i = 0; i < count; i++) args.push(randomValue(random, hostile));
      const label = `${name}(${describe(args)})`;
      let result: Value;
      try {
        result = spec.call(args, context);
      } catch (error) {
        assert.fail(`${label} threw ${String(error)}`);
      }
      assertValue(result, label);
    }
  }
}

test('the registry holds upper-case names and sound argument counts', () => {
  assert.ok(FUNCTIONS.size >= 170);
  for (const [name, spec] of FUNCTIONS) {
    assert.equal(name, name.toUpperCase());
    assert.equal(typeof spec.call, 'function');
    assert.ok(Number.isInteger(spec.min) && spec.min >= 0, name);
    assert.ok(spec.max === undefined || (Number.isInteger(spec.max) && spec.max >= spec.min), name);
  }
  // Special forms belong to the evaluator.
  for (const name of ['IF', 'IFERROR', 'AND', 'OR', 'NOT', 'ROW', 'COLUMN', 'OFFSET', 'INDIRECT', 'INDEX', 'CHOOSE', 'ISBLANK', 'N', 'T', 'LET']) {
    assert.equal(FUNCTIONS.has(name), false, name);
  }
});

test('fuzz: well-formed random arguments never throw and always give a valid value', () => {
  context = makeContext(fuzzSheet());
  fuzz(1, 250, false);
});

test('fuzz: the 1904 date system', () => {
  context = makeContext(fuzzSheet(), true);
  fuzz(2, 80, false);
});

test('fuzz: hostile arguments (NaN, Infinity, broken arrays and ranges, wrong argument counts)', () => {
  context = makeContext(fuzzSheet());
  fuzz(3, 250, true);
});

test('fuzz: a context that misbehaves', () => {
  const base = makeContext(fuzzSheet());
  context = {
    ...base,
    now: () => NaN,
    random: () => 7,
    usedExtent: () => ({ rows: -5, cols: NaN }),
    cellValue: (sheet, row, col) => (row % 3 === 0 ? (undefined as unknown as Scalar) : base.cellValue(sheet, row, col)),
    forEachCell(sheet, range, visit) {
      base.forEachCell(sheet, range, visit);
      // A cell outside the range that was asked for, and a hole.
      visit(1, range.bottom + 5, range.right + 5);
      visit(undefined as unknown as Scalar, range.top, range.left);
    },
  };
  fuzz(4, 60, false);
});
