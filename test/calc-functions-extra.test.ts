import assert from 'node:assert/strict';
import { test } from 'node:test';
import { EXTRA_FUNCTIONS } from '../src/calc/functions/extra';
import { dateSerial, MAX_ARRAY_CELLS, weekdayOf } from '../src/calc/functions/helpers';
import { calcError, isArray, isError, type CalcContext, type ErrorCode, type Scalar, type Value } from '../src/calc/values';
import { MAX_COLS, MAX_ROWS, parseRange } from '../src/model/sheet/types';

// ---------------------------------------------------------------------------
// A sheet made of a small map of cells, enough of a CalcContext to call functions directly.

type Cells = { [name: string]: Scalar };

/** Serial of 15 March 2024 12:00, the clock of every test context. */
const NOW = 45366.5;

function makeContext(cells: Cells, date1904 = false, random?: () => number): CalcContext {
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
    random:
      random ||
      (() => {
        seed = (seed * 1103515245 + 12345) % 2147483648;
        return seed / 2147483648;
      }),
  };
}

let context = makeContext({});

/** A reference: "A1", "A1:B3", or whole columns such as "A:C". */
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
  const spec = EXTRA_FUNCTIONS[name];
  assert.ok(spec, `${name} is exported`);
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

function date(year: number, month: number, day: number): number {
  return dateSerial(year, month, day, false);
}

// ---------------------------------------------------------------------------

test('every extra function has a spec with sane argument counts', () => {
  const names = Object.keys(EXTRA_FUNCTIONS);
  assert.equal(names.length, 48);
  for (const name of names) {
    const spec = EXTRA_FUNCTIONS[name];
    assert.equal(name, name.toUpperCase());
    assert.equal(typeof spec.call, 'function');
    assert.ok(spec.min >= 0 && (spec.max === undefined || spec.max >= spec.min), name);
  }
  // Arguments calculated as arrays even in an ordinary formula.
  const arrays = (name: string, count: number): boolean[] => {
    const selector = EXTRA_FUNCTIONS[name].arrayArgs;
    const out: boolean[] = [];
    for (let i = 0; i < count; i++) out.push(selector === true || (typeof selector === 'function' && selector(i)));
    return out;
  };
  assert.deepEqual(arrays('MMULT', 2), [true, true]);
  assert.deepEqual(arrays('FREQUENCY', 2), [true, true]);
  assert.deepEqual(arrays('FORECAST', 3), [false, true, true]);
  assert.deepEqual(arrays('FORECAST.LINEAR', 3), [false, true, true]);
  assert.deepEqual(arrays('TRIMMEAN', 2), [true, false]);
  assert.deepEqual(arrays('XNPV', 3), [false, true, true]);
  assert.deepEqual(arrays('XIRR', 3), [true, true, false]);
  assert.deepEqual(arrays('DSUM', 3), [false, false, false]);
  assert.equal(EXTRA_FUNCTIONS.SEQUENCE.scalarArgs, true);
  assert.equal(EXTRA_FUNCTIONS.RANDARRAY.scalarArgs, true);
});

test('HYPERLINK shows the friendly name, else the link', () => {
  check([
    ['https://example.com/report', 'HYPERLINK', 'https://example.com/report'],
    ['Report', 'HYPERLINK', 'https://example.com/report', 'Report'],
    [42, 'HYPERLINK', 'https://example.com', 42],
    ['#DIV/0!', 'HYPERLINK', 'https://example.com', err('#DIV/0!')],
    ['#N/A', 'HYPERLINK', err('#N/A')],
    [[['a', 'b']], 'HYPERLINK', 'https://example.com', arr([['a', 'b']])],
  ]);
  context = makeContext({ A1: 'https://example.com', B1: 'Click' });
  check([
    ['Click', 'HYPERLINK', ref('A1'), ref('B1')],
    ['https://example.com', 'HYPERLINK', ref('A1')],
  ]);
  context = makeContext({});
  // UNSURE: Excel may show the friendly name even when the link itself is an error; here too.
  check([['x', 'HYPERLINK', err('#N/A'), 'x']]);
});

test('ADDRESS: reference styles and kinds (Excel help examples)', () => {
  check([
    ['$C$2', 'ADDRESS', 2, 3],
    ['C$2', 'ADDRESS', 2, 3, 2],
    ['R2C[3]', 'ADDRESS', 2, 3, 2, false],
    ["'[Book1]Sheet1'!R2C3", 'ADDRESS', 2, 3, 1, false, '[Book1]Sheet1'],
    ["'EXCEL SHEET'!R2C3", 'ADDRESS', 2, 3, 1, false, 'EXCEL SHEET'],
    ['$C2', 'ADDRESS', 2, 3, 3],
    ['C2', 'ADDRESS', 2, 3, 4],
    ['R2C3', 'ADDRESS', 2, 3, 1, false],
    ['R[2]C3', 'ADDRESS', 2, 3, 3, false],
    ['R[2]C[3]', 'ADDRESS', 2, 3, 4, false],
    ['$A$1', 'ADDRESS', 1, 1, null, null],
    ['$A$1', 'ADDRESS', 1.9, 1.9, 1.9, true],
    ['$Z$1', 'ADDRESS', 1, 26],
    ['$AA$1', 'ADDRESS', 1, 27],
    ['$ZZ$1', 'ADDRESS', 1, 702],
    ['$AAA$1', 'ADDRESS', 1, 703],
    ['$XFD$1048576', 'ADDRESS', 1048576, 16384],
    ['Sheet1!$A$1', 'ADDRESS', 1, 1, 1, true, 'Sheet1'],
    ['Data_2024.v1!A1', 'ADDRESS', 1, 1, 4, true, 'Data_2024.v1'],
    ["'O''Neil'!$A$1", 'ADDRESS', 1, 1, 1, true, "O'Neil"],
    ["'Q1-Q2'!$A$1", 'ADDRESS', 1, 1, 1, true, 'Q1-Q2'],
    ['#VALUE!', 'ADDRESS', 0, 1],
    ['#VALUE!', 'ADDRESS', 1, 0],
    ['#VALUE!', 'ADDRESS', 1048577, 1],
    ['#VALUE!', 'ADDRESS', 1, 16385],
    ['#VALUE!', 'ADDRESS', 1, 1, 0],
    ['#VALUE!', 'ADDRESS', 1, 1, 5],
    ['#VALUE!', 'ADDRESS', 'x', 1],
    ['#N/A', 'ADDRESS', 1, err('#N/A')],
    [[['$A$1', '$B$1']], 'ADDRESS', 1, arr([[1, 2]])],
  ]);
  // UNSURE: whether Excel's ADDRESS quotes names that start with a digit, look like a cell or
  // hold symbols beyond ASCII. They are quoted here, which INDIRECT reads the same either way.
  check([
    ["'2024'!$A$1", 'ADDRESS', 1, 1, 1, true, '2024'],
    ["'A1'!$A$1", 'ADDRESS', 1, 1, 1, true, 'A1'],
    ["'RC'!$A$1", 'ADDRESS', 1, 1, 1, true, 'RC'],
    ['Özet!$A$1', 'ADDRESS', 1, 1, 1, true, 'Özet'],
    // UNSURE: an empty sheet text gives a bare "!".
    ['!$A$1', 'ADDRESS', 1, 1, 1, true, ''],
  ]);
  assert.equal(plain(call('ADDRESS', 1, 1, 1, true, 'x'.repeat(40000))), '#VALUE!');
});

test('CEILING.PRECISE, ISO.CEILING and FLOOR.PRECISE round up and down whatever the signs', () => {
  for (const name of ['CEILING.PRECISE', 'ISO.CEILING']) {
    check([
      [5, name, 4.3],
      [-4, name, -4.3],
      [6, name, 4.3, 2],
      [6, name, 4.3, -2],
      [-4, name, -4.3, 2],
      [-4, name, -4.3, -2],
      [0, name, 0, 5],
      [0, name, 4.3, 0],
      [4.4, name, 4.31, 0.1],
      [1.3, name, 1.3, 0.1],
      ['#VALUE!', name, 'x'],
    ]);
  }
  check([
    [-4, 'FLOOR.PRECISE', -3.2, -1],
    [3, 'FLOOR.PRECISE', 3.2, 1],
    [-4, 'FLOOR.PRECISE', -3.2, 1],
    [3, 'FLOOR.PRECISE', 3.2, -1],
    [3, 'FLOOR.PRECISE', 3.2],
    [-4, 'FLOOR.PRECISE', -3.2],
    [0, 'FLOOR.PRECISE', 3.2, 0],
    [1.3, 'FLOOR.PRECISE', 1.3, 0.1],
    [[[1, -2]], 'FLOOR.PRECISE', arr([[1.5, -1.5]])],
  ]);
});

test('ROMAN: the five forms (Excel help examples)', () => {
  check([
    ['CDXCIX', 'ROMAN', 499, 0],
    ['LDVLIV', 'ROMAN', 499, 1],
    ['XDIX', 'ROMAN', 499, 2],
    ['VDIV', 'ROMAN', 499, 3],
    ['ID', 'ROMAN', 499, 4],
    ['CMXCIX', 'ROMAN', 999, 0],
    ['LMVLIV', 'ROMAN', 999, 1],
    ['XMIX', 'ROMAN', 999, 2],
    ['VMIV', 'ROMAN', 999, 3],
    ['IM', 'ROMAN', 999, 4],
    ['CDXCIX', 'ROMAN', 499],
    ['CDXCIX', 'ROMAN', 499, true],
    ['ID', 'ROMAN', 499, false],
    ['', 'ROMAN', 0],
    ['I', 'ROMAN', 1],
    ['IV', 'ROMAN', 4],
    ['IX', 'ROMAN', 9],
    ['XIV', 'ROMAN', 14],
    ['XL', 'ROMAN', 40],
    ['XC', 'ROMAN', 90],
    ['CD', 'ROMAN', 400],
    ['MCMXCIV', 'ROMAN', 1994],
    ['MMXXIV', 'ROMAN', 2024.9],
    ['MMMCMXCIX', 'ROMAN', 3999],
    ['MMMDCCCLXXXVIII', 'ROMAN', 3888, 4],
    ['#VALUE!', 'ROMAN', 4000],
    ['#VALUE!', 'ROMAN', -1],
    ['#VALUE!', 'ROMAN', 10, 5],
    ['#VALUE!', 'ROMAN', 10, -1],
    ['#VALUE!', 'ROMAN', 'x'],
  ]);
  // Whatever the form, the numerals add up to the number again.
  const worth: { [numeral: string]: number } = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
  for (let form = 0; form <= 4; form++) {
    for (let n = 1; n <= 3999; n++) {
      const text = call('ROMAN', n, form) as string;
      let total = 0;
      for (let i = 0; i < text.length; i++) {
        const value = worth[text.charAt(i)];
        const next = i + 1 < text.length ? worth[text.charAt(i + 1)] : 0;
        total += value < next ? -value : value;
      }
      assert.equal(total, n, `ROMAN(${n}, ${form}) = ${text}`);
    }
  }
});

test('SEQUENCE', () => {
  check([
    [[[1], [2], [3]], 'SEQUENCE', 3],
    [[[1, 2, 3], [4, 5, 6]], 'SEQUENCE', 2, 3],
    [[[10, 8], [6, 4]], 'SEQUENCE', 2, 2, 10, -2],
    [[[5], [6]], 'SEQUENCE', 2, null, 5],
    [[[0, 0.5, 1]], 'SEQUENCE', 1, 3, 0, 0.5],
    [[[1, 2]], 'SEQUENCE', 1.9, 2.9],
    ['#CALC!', 'SEQUENCE', 0],
    ['#CALC!', 'SEQUENCE', 2, 0],
    ['#VALUE!', 'SEQUENCE', -1],
    ['#VALUE!', 'SEQUENCE', 'x'],
    ['#N/A', 'SEQUENCE', 2, 2, err('#N/A')],
    ['#NUM!', 'SEQUENCE', 4096, 4096],
    ['#NUM!', 'SEQUENCE', 1e200, 1e200],
    ['#NUM!', 'SEQUENCE', 2, 1, 1e308, 1e308],
  ]);
  const largest = call('SEQUENCE', MAX_ARRAY_CELLS / 2, 2);
  assert.ok(isArray(largest) && largest.data.length === MAX_ARRAY_CELLS && largest.data[MAX_ARRAY_CELLS - 1] === MAX_ARRAY_CELLS);
});

test('RANDARRAY', () => {
  const draws = [0, 0.5, 0.999999];
  let at = 0;
  context = makeContext({}, false, () => draws[at++ % draws.length]);
  check([
    [[[0]], 'RANDARRAY'],
    [[[0.5, 0.999999]], 'RANDARRAY', 1, 2],
    [[[10], [15], [19.99999]], 'RANDARRAY', 3, 1, 10, 20],
    [[[1, 4, 6]], 'RANDARRAY', 1, 3, 1, 6, true],
    [[[-2, 0, 2]], 'RANDARRAY', 1, 3, -2, 2, true],
    [[[7, 7, 7]], 'RANDARRAY', 1, 3, 7, 7, true],
    ['#VALUE!', 'RANDARRAY', 1, 1, 5, 1],
    ['#CALC!', 'RANDARRAY', 0],
    ['#VALUE!', 'RANDARRAY', -1],
    ['#NUM!', 'RANDARRAY', 4096, 4096],
    // UNSURE: Excel is taken to refuse limits that are not whole when whole numbers are asked for.
    ['#VALUE!', 'RANDARRAY', 1, 1, 0.5, 3, true],
  ]);
  // A generator that misbehaves cannot push a value outside the limits.
  const bad = [NaN, 1, -1, Infinity, 2];
  context = makeContext({}, false, () => bad[at++ % bad.length]);
  const result = call('RANDARRAY', 5, 4, 1, 6, true);
  assert.ok(isArray(result));
  for (const value of result.data) assert.ok(typeof value === 'number' && value >= 1 && value <= 6 && Math.floor(value) === value);
  context = makeContext({});
  const many = call('RANDARRAY', 100, 100, 3, 4);
  assert.ok(isArray(many) && many.rows === 100 && many.cols === 100);
  for (const value of many.data) assert.ok(typeof value === 'number' && value >= 3 && value < 4);
});

test('MMULT', () => {
  check([
    [[[2, 6], [14, 4]], 'MMULT', arr([[1, 3], [7, 2]]), arr([[2, 0], [0, 2]])],
    [[[11]], 'MMULT', arr([[1, 2]]), arr([[3], [4]])],
    [[[3, 6], [4, 8]], 'MMULT', arr([[3], [4]]), arr([[1, 2]])],
    [[[58, 64], [139, 154]], 'MMULT', arr([[1, 2, 3], [4, 5, 6]]), arr([[7, 8], [9, 10], [11, 12]])],
    [[[6]], 'MMULT', 2, 3],
    ['#VALUE!', 'MMULT', arr([[1, 2]]), arr([[1, 2]])],
    ['#VALUE!', 'MMULT', arr([[1, 'x']]), arr([[3], [4]])],
    ['#VALUE!', 'MMULT', arr([[1, null]]), arr([[3], [4]])],
    ['#VALUE!', 'MMULT', arr([[1, true]]), arr([[3], [4]])],
    ['#DIV/0!', 'MMULT', arr([[1, err('#DIV/0!')]]), arr([[3], [4]])],
    [[['#NUM!']], 'MMULT', arr([[1e200]]), arr([[1e200]])],
  ]);
  context = makeContext({ A1: 1, B1: 2, A2: 3, B2: 4, D1: 5, D2: 6 });
  check([
    [[[17], [39]], 'MMULT', ref('A1:B2'), ref('D1:D2')],
    ['#VALUE!', 'MMULT', ref('A1:B3'), ref('D1:D2')],
  ]);
  context = makeContext({});
  // A product that would take too long is refused instead of being worked out.
  const wide: Value = { rows: 1024, cols: 1024, data: new Array<Scalar>(1024 * 1024).fill(1) };
  assert.equal(plain(call('MMULT', wide, wide)), '#NUM!');
});

// ---------------------------------------------------------------------------
// Working days

/** The days off of a weekend code, by weekday with 0 for Sunday, as the Excel help lists them. */
const WEEKENDS: { [code: string]: number[] } = {
  1: [6, 0],
  2: [0, 1],
  3: [1, 2],
  4: [2, 3],
  5: [3, 4],
  6: [4, 5],
  7: [5, 6],
  11: [0],
  12: [1],
  13: [2],
  14: [3],
  15: [4],
  16: [5],
  17: [6],
};

function slowNetworkDays(start: number, end: number, off: number[], holidays: number[]): number {
  let count = 0;
  for (let day = Math.min(start, end); day <= Math.max(start, end); day++) {
    if (off.indexOf(weekdayOf(day, false)) < 0 && holidays.indexOf(day) < 0) count++;
  }
  return start > end ? -count : count;
}

function slowWorkDay(start: number, count: number, off: number[], holidays: number[]): number {
  let day = start;
  const step = count < 0 ? -1 : 1;
  for (let left = Math.abs(count); left > 0; ) {
    day += step;
    if (off.indexOf(weekdayOf(day, false)) < 0 && holidays.indexOf(day) < 0) left--;
  }
  return day;
}

test('NETWORKDAYS.INTL (Excel help examples)', () => {
  const holidays = arr([['2006/1/2', '2006/1/16']]);
  check([
    [22, 'NETWORKDAYS.INTL', date(2006, 1, 1), date(2006, 1, 31)],
    [-21, 'NETWORKDAYS.INTL', date(2006, 2, 28), date(2006, 1, 31)],
    [22, 'NETWORKDAYS.INTL', date(2006, 1, 1), date(2006, 2, 1), 7, holidays],
    [20, 'NETWORKDAYS.INTL', date(2006, 1, 1), date(2006, 2, 1), '0010001', holidays],
    [22, 'NETWORKDAYS.INTL', date(2006, 1, 1), date(2006, 1, 31), null],
    [22, 'NETWORKDAYS.INTL', date(2006, 1, 1), date(2006, 1, 31), 1],
    [22, 'NETWORKDAYS.INTL', date(2006, 1, 1), date(2006, 1, 31), '0000011'],
    [31, 'NETWORKDAYS.INTL', date(2006, 1, 1), date(2006, 1, 31), '0000000'],
    [1, 'NETWORKDAYS.INTL', date(2006, 1, 2), date(2006, 1, 2)],
    [0, 'NETWORKDAYS.INTL', date(2006, 1, 1), date(2006, 1, 1)],
    ['#VALUE!', 'NETWORKDAYS.INTL', date(2006, 1, 1), date(2006, 1, 31), '1111111'],
    ['#VALUE!', 'NETWORKDAYS.INTL', date(2006, 1, 1), date(2006, 1, 31), '000001'],
    ['#VALUE!', 'NETWORKDAYS.INTL', date(2006, 1, 1), date(2006, 1, 31), '0000012'],
    ['#NUM!', 'NETWORKDAYS.INTL', date(2006, 1, 1), date(2006, 1, 31), 0],
    ['#NUM!', 'NETWORKDAYS.INTL', date(2006, 1, 1), date(2006, 1, 31), 8],
    ['#NUM!', 'NETWORKDAYS.INTL', date(2006, 1, 1), date(2006, 1, 31), 10],
    ['#NUM!', 'NETWORKDAYS.INTL', date(2006, 1, 1), date(2006, 1, 31), 18],
    ['#NUM!', 'NETWORKDAYS.INTL', -1, date(2006, 1, 31)],
    ['#VALUE!', 'NETWORKDAYS.INTL', 'x', date(2006, 1, 31)],
    ['#VALUE!', 'NETWORKDAYS.INTL', date(2006, 1, 1), date(2006, 1, 31), 1, arr([['not a date']])],
    ['#N/A', 'NETWORKDAYS.INTL', date(2006, 1, 1), date(2006, 1, 31), 1, arr([[err('#N/A')]])],
  ]);
  // Monday 2 January 2006 is a holiday given once as a number and once as text; it counts once.
  context = makeContext({ A1: date(2006, 1, 2), A2: '2006/1/2', A3: date(2006, 1, 7) });
  check([[21, 'NETWORKDAYS.INTL', date(2006, 1, 1), date(2006, 1, 31), 1, ref('A1:A5')]]);
  context = makeContext({});
});

test('NETWORKDAYS.INTL and WORKDAY.INTL agree with counting day by day', () => {
  const base = date(2024, 2, 20);
  const holidays = [base + 3, base + 4, base + 5, base + 9, base + 30, base - 2, base - 3, base - 20, base + 3];
  const list = arr([holidays]);
  for (const code of Object.keys(WEEKENDS)) {
    const off = WEEKENDS[code];
    for (let start = base; start < base + 7; start++) {
      for (let span = -45; span <= 45; span += span > -9 && span < 9 ? 1 : 6) {
        const label = `weekend ${code}, start ${start - base}, ${span}`;
        assert.equal(call('NETWORKDAYS.INTL', start, start + span, Number(code), list), slowNetworkDays(start, start + span, off, holidays), `net ${label}`);
        assert.equal(call('NETWORKDAYS.INTL', start, start + span, Number(code)), slowNetworkDays(start, start + span, off, []), `net plain ${label}`);
        if (span === 0) continue;
        assert.equal(call('WORKDAY.INTL', start, span, Number(code), list), slowWorkDay(start, span, off, holidays), `work ${label}`);
        assert.equal(call('WORKDAY.INTL', start, span, Number(code)), slowWorkDay(start, span, off, []), `work plain ${label}`);
      }
    }
  }
  // A weekend written as text: only Wednesday and Saturday are worked.
  const off = [1, 2, 4, 5, 0];
  for (let span = -30; span <= 30; span++) {
    assert.equal(call('NETWORKDAYS.INTL', base, base + span, '1101101', list), slowNetworkDays(base, base + span, off, holidays));
    if (span !== 0) assert.equal(call('WORKDAY.INTL', base, span, '1101101', list), slowWorkDay(base, span, off, holidays));
  }
});

test('WORKDAY.INTL (Excel help examples)', () => {
  check([
    ['#NUM!', 'WORKDAY.INTL', date(2012, 1, 1), 30, 0],
    [41013, 'WORKDAY.INTL', date(2012, 1, 1), 90, 11],
    [date(2012, 2, 5), 'WORKDAY.INTL', date(2012, 1, 1), 30, 17],
    // 1 January 2012 is a Sunday: no days to move leaves the date as it is.
    [date(2012, 1, 1), 'WORKDAY.INTL', date(2012, 1, 1), 0],
    [date(2012, 1, 2), 'WORKDAY.INTL', date(2012, 1, 1), 1],
    [date(2011, 12, 30), 'WORKDAY.INTL', date(2012, 1, 1), -1],
    [date(2012, 1, 3), 'WORKDAY.INTL', date(2012, 1, 1), 1, 1, arr([[date(2012, 1, 2)]])],
    [date(2012, 1, 3), 'WORKDAY.INTL', date(2012, 1, 1.5), 1.9, 1, arr([['2012/1/2']])],
    ['#VALUE!', 'WORKDAY.INTL', date(2012, 1, 1), 1, '1111111'],
    ['#VALUE!', 'WORKDAY.INTL', date(2012, 1, 1), 1, 'weekend'],
    ['#NUM!', 'WORKDAY.INTL', date(2012, 1, 1), 1, 9],
    ['#NUM!', 'WORKDAY.INTL', date(2012, 1, 1), -100000],
    ['#NUM!', 'WORKDAY.INTL', date(2012, 1, 1), 5000000],
    ['#NUM!', 'WORKDAY.INTL', date(2012, 1, 1), 1e300],
    ['#VALUE!', 'WORKDAY.INTL', date(2012, 1, 1), 'x'],
  ]);
  // A long way is not walked day by day.
  const far = call('WORKDAY.INTL', 2, 2000000, 11) as number;
  assert.equal(far, slowWorkDay(2, 2000000, [0], []));
});

// ---------------------------------------------------------------------------
// Engineering

test('DEC2BIN, DEC2OCT, DEC2HEX (Excel help examples and limits)', () => {
  check([
    ['1001', 'DEC2BIN', 9, 4],
    ['1110011100', 'DEC2BIN', -100],
    ['0', 'DEC2BIN', 0],
    ['111111111', 'DEC2BIN', 511],
    ['1000000000', 'DEC2BIN', -512],
    ['1111111111', 'DEC2BIN', -1],
    ['1111111111', 'DEC2BIN', -1, 3],
    ['101', 'DEC2BIN', 5.9],
    ['0000000101', 'DEC2BIN', 5, 10],
    ['101', 'DEC2BIN', '5'],
    ['#NUM!', 'DEC2BIN', 512],
    ['#NUM!', 'DEC2BIN', -513],
    ['#NUM!', 'DEC2BIN', 9, 3],
    ['#NUM!', 'DEC2BIN', 9, 0],
    ['#NUM!', 'DEC2BIN', 9, -1],
    ['#NUM!', 'DEC2BIN', 9, 11],
    ['#VALUE!', 'DEC2BIN', 'x'],
    ['#VALUE!', 'DEC2BIN', 9, 'x'],
    ['#VALUE!', 'DEC2BIN', true],
    ['#N/A', 'DEC2BIN', err('#N/A')],
    ['0064', 'DEC2HEX', 100, 4],
    ['FFFFFFFFCA', 'DEC2HEX', -54],
    ['1C', 'DEC2HEX', 28],
    ['7FFFFFFFFF', 'DEC2HEX', 549755813887],
    ['8000000000', 'DEC2HEX', -549755813888],
    ['#NUM!', 'DEC2HEX', 549755813888],
    ['#NUM!', 'DEC2HEX', -549755813889],
    ['#NUM!', 'DEC2HEX', 64, 1],
    ['072', 'DEC2OCT', 58, 3],
    ['7777777634', 'DEC2OCT', -100],
    ['3777777777', 'DEC2OCT', 536870911],
    ['4000000000', 'DEC2OCT', -536870912],
    ['#NUM!', 'DEC2OCT', 536870912],
    ['#NUM!', 'DEC2OCT', -536870913],
    [[['1', '10', '11']], 'DEC2BIN', arr([[1, 2, 3]])],
  ]);
});

test('BIN2DEC, OCT2DEC, HEX2DEC (Excel help examples and limits)', () => {
  check([
    [100, 'BIN2DEC', '1100100'],
    [-1, 'BIN2DEC', '1111111111'],
    [100, 'BIN2DEC', 1100100],
    [511, 'BIN2DEC', '0111111111'],
    [-512, 'BIN2DEC', '1000000000'],
    ['#NUM!', 'BIN2DEC', '11111111111'],
    ['#NUM!', 'BIN2DEC', '102'],
    ['#NUM!', 'BIN2DEC', '-1'],
    ['#NUM!', 'BIN2DEC', 1.5],
    ['#NUM!', 'BIN2DEC', ' 1'],
    ['#VALUE!', 'BIN2DEC', true],
    ['#N/A', 'BIN2DEC', err('#N/A')],
    [165, 'HEX2DEC', 'A5'],
    [-165, 'HEX2DEC', 'FFFFFFFF5B'],
    [1034160313, 'HEX2DEC', '3DA408B9'],
    [255, 'HEX2DEC', 'ff'],
    [549755813887, 'HEX2DEC', '7FFFFFFFFF'],
    [-549755813888, 'HEX2DEC', '8000000000'],
    ['#NUM!', 'HEX2DEC', 'G'],
    ['#NUM!', 'HEX2DEC', '12345678901'],
    [44, 'OCT2DEC', 54],
    [-165, 'OCT2DEC', '7777777533'],
    [536870911, 'OCT2DEC', '3777777777'],
    [-536870912, 'OCT2DEC', '4000000000'],
    ['#NUM!', 'OCT2DEC', '8'],
    ['#NUM!', 'HEX2DEC', 'F'.repeat(40000)],
    // UNSURE: empty text is taken as 0, as Excel is believed to do.
    [0, 'BIN2DEC', ''],
    [0, 'HEX2DEC', null],
  ]);
  // Every number comes back from its own digits.
  for (const [to, from, low, high] of [
    ['DEC2BIN', 'BIN2DEC', -512, 511],
    ['DEC2OCT', 'OCT2DEC', -536870912, 536870911],
    ['DEC2HEX', 'HEX2DEC', -549755813888, 549755813887],
  ] as [string, string, number, number][]) {
    for (const n of [low, low + 1, -2, -1, 0, 1, 2, 7, 8, 15, 16, 255, high - 1, high]) {
      assert.equal(call(from, call(to, n)), n, `${from}(${to}(${n}))`);
    }
  }
});

test('BITAND, BITOR, BITXOR (Excel help examples, 48 bits)', () => {
  check([
    [1, 'BITAND', 1, 5],
    [9, 'BITAND', 13, 25],
    [31, 'BITOR', 23, 10],
    [6, 'BITXOR', 5, 3],
    [0, 'BITAND', 0, 0],
    // Above 2^31, where the operators of the language would wrap around.
    [2147483648, 'BITAND', 2147483648, 4294967295],
    [4294967295, 'BITOR', 2147483648, 2147483647],
    [281474976710655, 'BITOR', 281474976710655, 1],
    [281470681743359, 'BITXOR', 281474976710655, 4294967296],
    [281474976710654, 'BITXOR', 281474976710655, 1],
    [4294967296, 'BITAND', 281474976710655, 4294967296],
    [0, 'BITAND', 140737488355328, 140737488355327],
    [281474976710655, 'BITOR', 140737488355328, 140737488355327],
    ['#NUM!', 'BITAND', 281474976710656, 1],
    ['#NUM!', 'BITOR', -1, 1],
    ['#NUM!', 'BITXOR', 1.5, 1],
    ['#VALUE!', 'BITAND', 'x', 1],
    ['#N/A', 'BITAND', 1, err('#N/A')],
  ]);
});

// ---------------------------------------------------------------------------
// Database functions (the table of the Excel help)

const ORCHARD: Cells = {
  A1: 'Tree', B1: 'Height', C1: 'Age', D1: 'Yield', E1: 'Profit', F1: 'Height',
  A2: '=Apple', B2: '>10', F2: '<16',
  A3: '=Pear',
  A5: 'Tree', B5: 'Height', C5: 'Age', D5: 'Yield', E5: 'Profit',
  A6: 'Apple', B6: 18, C6: 20, D6: 14, E6: 105,
  A7: 'Pear', B7: 12, C7: 12, D7: 10, E7: 96,
  A8: 'Cherry', B8: 13, C8: 14, D8: 9, E8: 105,
  A9: 'Apple', B9: 14, C9: 15, D9: 10, E9: 75,
  A10: 'Pear', B10: 9, C10: 8, D10: 8, E10: 77,
  A11: 'Apple', B11: 8, C11: 9, D11: 6, E11: 45,
};

test('database functions (Excel help examples)', () => {
  context = makeContext(ORCHARD);
  const db = ref('A5:E11');
  check([
    [225, 'DSUM', db, 'Profit', ref('A1:A2')],
    [248, 'DSUM', db, 'Profit', ref('A1:F3')],
    [1, 'DCOUNT', db, 'Age', ref('A1:F2')],
    [1, 'DCOUNTA', db, 'Profit', ref('A1:F2')],
    [96, 'DMAX', db, 'Profit', ref('A1:F3')],
    [75, 'DMIN', db, 'Profit', ref('A1:B2')],
    [12, 'DAVERAGE', db, 'Yield', ref('A1:B2')],
    // The database as its own criteria: every record meets the row that describes it.
    [13, 'DAVERAGE', db, 3, db],
    ['#NUM!', 'DGET', db, 'Yield', ref('A1:A3')],
    [10, 'DGET', db, 'Yield', ref('A1:F2')],
  ]);
  context = makeContext({});
});

test('database functions: fields, criteria and what they count', () => {
  context = makeContext({
    ...ORCHARD,
    H1: 'Tree', H2: 'app', // text without an operator: begins with
    I1: 'tree', I2: '=app', // exact, and no tree is called that
    J1: 'Height', J2: 13, // a number
    K1: 'Age', K2: '>=14', L1: 'Yield', L2: '<10', // two columns of one row
    M1: 'Tree', M2: 'Ch*y', M3: '<>Apple',
    N1: 'Tree', // a criteria row that asks for nothing
    O1: 'Color', O2: 'red', // a field the database does not have
    P1: 'Color',
    Q1: 'Tree', Q2: 'Plum',
    A12: 'Fig', B12: 'tall', D12: err('#DIV/0!'), E12: true,
    A13: null,
  });
  const db = ref('A5:E12');
  check([
    // UNSURE: plain text is taken to match what begins with it, as the advanced filter does.
    [225, 'DSUM', db, 'Profit', ref('H1:H2')],
    [0, 'DSUM', db, 'Profit', ref('I1:I2')],
    [105, 'DSUM', db, 'profit', ref('J1:J2')],
    [105, 'DSUM', db, 5, ref('K1:L2')],
    [105, 'DSUM', db, 5.9, ref('K1:L2')],
    // Cherry by the pattern, then everything that is not an apple (Cherry once, the pears, and a fig without a number).
    [278, 'DSUM', db, 'Profit', ref('M1:M3')],
    [3, 'DCOUNT', db, 'Profit', ref('M1:M3')],
    [4, 'DCOUNTA', db, 'Profit', ref('M1:M3')],
    // No criteria at all: every record.
    [503, 'DSUM', db, 'Profit', ref('N1:N2')],
    [6, 'DCOUNT', db, 'Height', ref('N1:N2')],
    [7, 'DCOUNTA', db, 'Height', ref('N1:N2')],
    [7, 'DCOUNT', db, null, ref('N1:N2')],
    [7, 'DCOUNTA', db, null, ref('N1:N2')],
    [3, 'DCOUNT', db, null, ref('H1:H2')],
    [105, 'DMAX', db, 'Profit', ref('N1:N2')],
    [45, 'DMIN', db, 'Profit', ref('N1:N2')],
    [6, 'DMIN', db, 'Yield', ref('H1:H2')],
    ['#DIV/0!', 'DSUM', db, 'Yield', ref('N1:N2')],
    ['#DIV/0!', 'DAVERAGE', db, 'Yield', ref('N1:N2')],
    [6, 'DCOUNT', db, 'Yield', ref('N1:N2')],
    [7, 'DCOUNTA', db, 'Yield', ref('N1:N2')],
    // Nothing matches.
    [0, 'DSUM', db, 'Profit', ref('Q1:Q2')],
    [0, 'DCOUNT', db, 'Profit', ref('Q1:Q2')],
    [0, 'DMAX', db, 'Profit', ref('Q1:Q2')],
    [0, 'DMIN', db, 'Profit', ref('Q1:Q2')],
    ['#DIV/0!', 'DAVERAGE', db, 'Profit', ref('Q1:Q2')],
    ['#VALUE!', 'DGET', db, 'Profit', ref('Q1:Q2')],
    [105, 'DGET', db, 'Profit', ref('M1:M2')],
    ['Cherry', 'DGET', db, 1, ref('M1:M2')],
    // Fields that do not exist.
    ['#VALUE!', 'DSUM', db, 'Color', ref('N1:N2')],
    ['#VALUE!', 'DSUM', db, 0, ref('N1:N2')],
    ['#VALUE!', 'DSUM', db, 6, ref('N1:N2')],
    ['#VALUE!', 'DSUM', db, true, ref('N1:N2')],
    ['#VALUE!', 'DSUM', db, null, ref('N1:N2')],
    ['#N/A', 'DSUM', db, err('#N/A'), ref('N1:N2')],
    // A criteria heading without a field is ignored while nothing stands under it ...
    [503, 'DSUM', db, 'Profit', ref('P1:P2')],
    // ... and refused when something does: UNSURE what Excel answers there, so no number is made up.
    ['#VALUE!', 'DSUM', db, 'Profit', ref('O1:O2')],
    // Criteria need a heading row and at least one row below it.
    ['#VALUE!', 'DSUM', db, 'Profit', ref('H1:H1')],
    ['#REF!', 'DSUM', err('#REF!'), 'Profit', ref('H1:H2')],
    ['#REF!', 'DSUM', db, 'Profit', err('#REF!')],
  ]);
  // Arrays work as well as references.
  const table = arr([
    ['Name', 'Qty'],
    ['a', 1],
    ['b', 2],
    ['a', 4],
  ]);
  check([
    [5, 'DSUM', table, 'Qty', arr([['Name'], ['=a']])],
    [3, 'DCOUNT', table, 2, arr([['Qty'], ['>0']])],
    [6, 'DSUM', table, 'Qty', arr([['Name', 'Qty'], ['=a', '>1'], ['=b', null]])],
    [2, 'DGET', table, 'Qty', arr([['Name'], ['b']])],
  ]);
  context = makeContext({});
});

test('database functions on whole columns read only the used rows', () => {
  context = makeContext({ A1: 'Name', B1: 'Qty', A2: 'a', B2: 1, A3: 'b', B3: 2, A4: 'a', B4: 4, D1: 'Name', D2: '=a' });
  check([
    [5, 'DSUM', ref('A:B'), 'Qty', ref('D1:D2')],
    [2, 'DCOUNT', ref('A:B'), 'Qty', ref('D1:D2')],
    [4, 'DMAX', ref('A:B'), 2, ref('D1:D2')],
    ['#NUM!', 'DGET', ref('A:B'), 'Qty', ref('D1:D2')],
  ]);
  context = makeContext({});
});

// ---------------------------------------------------------------------------

test('nothing throws and no NaN or Infinity comes back, whatever the arguments', () => {
  context = makeContext({ A1: 'Tree', A2: 'x', B1: 1e308, B2: -1e308, C1: err('#REF!'), C3: true });
  const samples: Value[] = [
    null,
    0,
    1,
    -1,
    0.5,
    1e308,
    -1e308,
    NaN,
    Infinity,
    -Infinity,
    '',
    'x',
    '0000011',
    'x'.repeat(40000),
    true,
    false,
    err('#N/A'),
    arr([[1, 2], [3, 4]]),
    arr([['a', null, err('#DIV/0!')]]),
    { rows: 0, cols: 0, data: [] },
    { rows: 3, cols: 3, data: [1] },
    ref('A1:C3'),
    ref('A:C'),
    { ref: true, sheet: 0, range: { top: 5, left: 5, bottom: 1, right: 1 } },
    { ref: true, sheet: 9, range: { top: 0, left: 0, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 } },
  ];
  const sound = (value: Value, label: string): void => {
    if (typeof value === 'number') assert.ok(isFinite(value), `${label} gave ${value}`);
    else if (isArray(value)) {
      for (const item of value.data) assert.ok(typeof item !== 'number' || isFinite(item), `${label} gave ${item} inside an array`);
    }
  };
  let seed = 7;
  const pick = (): Value => {
    seed = (seed * 1103515245 + 12345) % 2147483648;
    return samples[seed % samples.length];
  };
  for (const name of Object.keys(EXTRA_FUNCTIONS)) {
    const spec = EXTRA_FUNCTIONS[name];
    const most = spec.max === undefined ? spec.min + 2 : spec.max;
    for (let count = spec.min; count <= most; count++) {
      const rounds = count === 0 ? 1 : 150;
      for (let round = 0; round < rounds; round++) {
        const args: Value[] = [];
        for (let i = 0; i < count; i++) args.push(pick());
        const label = `${name}(${args.map((arg) => JSON.stringify(plain(arg)).slice(0, 30)).join(', ')})`;
        let result: Value;
        try {
          result = spec.call(args, context);
        } catch (error) {
          assert.fail(`${label} threw ${String(error)}`);
        }
        assert.notEqual(result, undefined, label);
        sound(result, label);
      }
    }
  }
  context = makeContext({});
});
