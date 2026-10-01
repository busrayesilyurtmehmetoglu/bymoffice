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

/** Serial of a date in the 1900 system (valid from 1 March 1900), computed independently of the engine. */
function d(year: number, month: number, day: number): number {
  return Math.round(Date.UTC(year, month - 1, day) / 86400000) + 25569;
}

function hms(hour: number, minute: number, second: number): number {
  return (hour * 3600 + minute * 60 + second) / 86400;
}

test('DATE builds serials, carries overflowing months and days, and keeps the 1900 leap year bug', () => {
  context = makeContext({});
  check([
    [45366, 'DATE', 2024, 3, 15],
    [d(2009, 2, 2), 'DATE', 2008, 14, 2],
    [d(2007, 9, 2), 'DATE', 2008, -3, 2],
    [d(2008, 2, 4), 'DATE', 2008, 1, 35],
    [d(2007, 12, 16), 'DATE', 2008, 1, -15],
    [d(2007, 12, 31), 'DATE', 2008, 1, 0],
    [d(2007, 12, 1), 'DATE', 2008, 0, 1],
    // Years below 1900 are counted from 1900.
    [d(2008, 1, 2), 'DATE', 108, 1, 2],
    [d(3799, 1, 1), 'DATE', 1899, 1, 1],
    [1, 'DATE', 1900, 1, 1],
    [1, 'DATE', 0, 1, 1],
    [59, 'DATE', 1900, 2, 28],
    // 29 February 1900 never was, but it has a serial.
    [60, 'DATE', 1900, 2, 29],
    [60, 'DATE', 1900, 3, 0],
    [61, 'DATE', 1900, 3, 1],
    [0, 'DATE', 1900, 1, 0],
    [2958465, 'DATE', 9999, 12, 31],
    ['#NUM!', 'DATE', 9999, 12, 32],
    ['#NUM!', 'DATE', 10000, 1, 1],
    ['#NUM!', 'DATE', -1, 1, 1],
    ['#NUM!', 'DATE', 1900, 1, -1],
    ['#NUM!', 'DATE', 2024, 1e300, 1],
    ['#VALUE!', 'DATE', 'x', 1, 1],
    [45366, 'DATE', 2024.9, 3.9, 15.9],
    [[[d(2024, 1, 1), d(2024, 2, 1)]], 'DATE', 2024, arr([[1, 2]]), 1],
  ]);
  context = makeContext({}, true);
  check([
    [0, 'DATE', 1904, 1, 1],
    [1, 'DATE', 1904, 1, 2],
    [45366 - 1462, 'DATE', 2024, 3, 15],
    ['#NUM!', 'DATE', 1903, 12, 31],
  ]);
});

test('parts of a date and a time', () => {
  context = makeContext({});
  check([
    [2024, 'YEAR', 45366],
    [3, 'MONTH', 45366],
    [15, 'DAY', 45366],
    [2024, 'YEAR', 45366.999],
    [15, 'DAY', 45366.999],
    [1900, 'YEAR', 0],
    [1, 'MONTH', 0],
    [0, 'DAY', 0],
    [29, 'DAY', 60],
    [2, 'MONTH', 60],
    [1, 'DAY', 61],
    [9999, 'YEAR', 2958465],
    ['#NUM!', 'YEAR', 2958466],
    ['#NUM!', 'YEAR', -1],
    ['#VALUE!', 'YEAR', 'soon'],
    [2024, 'YEAR', '3/15/2024'],
    [2024, 'YEAR', '2024-03-15'],
    [1900, 'YEAR', true],
    [[[2024, 2025]], 'YEAR', arr([[45366, 45658]])],
    [18, 'HOUR', 0.75],
    [45, 'MINUTE', hms(12, 45, 0)],
    [18, 'SECOND', hms(16, 48, 18)],
    [16, 'HOUR', 45366 + hms(16, 48, 18)],
    [48, 'MINUTE', 45366 + hms(16, 48, 18)],
    [0, 'HOUR', 45366],
    [18, 'HOUR', '6:30 PM'],
    [30, 'MINUTE', '6:30 PM'],
    ['#NUM!', 'HOUR', -0.5],
    [0.5, 'TIME', 12, 0, 0],
    [hms(16, 48, 10), 'TIME', 16, 48, 10],
    // Hours past 24 wrap around; minutes and seconds carry.
    [hms(1, 0, 0), 'TIME', 25, 0, 0],
    [hms(1, 1, 5), 'TIME', 0, 60, 65],
    [hms(0, 30, 0), 'TIME', 1, -30, 0],
    ['#NUM!', 'TIME', 0, -1, 0],
    ['#NUM!', 'TIME', 32768, 0, 0],
    [NOW, 'NOW'],
    [45366, 'TODAY'],
  ]);
  // UNSURE: a time within half a second of midnight is read as 0:00:00 (of the same day).
  check([
    [0, 'HOUR', 0.999999],
    [0, 'SECOND', 0.999999],
  ]);
  context = makeContext({}, true);
  check([
    [1904, 'YEAR', 0],
    [1, 'DAY', 0],
    [2024, 'YEAR', 45366 - 1462],
  ]);
});

test('WEEKDAY, WEEKNUM and ISOWEEKNUM', () => {
  context = makeContext({});
  const thursday = d(2008, 2, 14);
  check([
    [5, 'WEEKDAY', thursday],
    [4, 'WEEKDAY', thursday, 2],
    [3, 'WEEKDAY', thursday, 3],
    [5, 'WEEKDAY', thursday, 1],
    [4, 'WEEKDAY', thursday, 11],
    [3, 'WEEKDAY', thursday, 12],
    [2, 'WEEKDAY', thursday, 13],
    [1, 'WEEKDAY', thursday, 14],
    [7, 'WEEKDAY', thursday, 15],
    [6, 'WEEKDAY', thursday, 16],
    [5, 'WEEKDAY', thursday, 17],
    ['#NUM!', 'WEEKDAY', thursday, 4],
    ['#NUM!', 'WEEKDAY', thursday, 18],
    [1, 'WEEKDAY', d(2024, 3, 17)],
    [7, 'WEEKDAY', d(2024, 3, 17), 2],
    [6, 'WEEKDAY', d(2024, 3, 17), 3],
    [7, 'WEEKDAY', d(2024, 3, 16)],
    // The 1900 system shows serial 1 as a Sunday.
    [1, 'WEEKDAY', 1],
    [10, 'WEEKNUM', d(2012, 3, 9)],
    [11, 'WEEKNUM', d(2012, 3, 9), 2],
    [1, 'WEEKNUM', d(2024, 1, 1)],
    [1, 'WEEKNUM', d(2024, 1, 6)],
    [2, 'WEEKNUM', d(2024, 1, 7)],
    [1, 'WEEKNUM', d(2024, 1, 7), 2],
    [2, 'WEEKNUM', d(2024, 1, 8), 2],
    [2, 'WEEKNUM', d(2024, 1, 8), 11],
    [53, 'WEEKNUM', d(2024, 12, 31)],
    [1, 'WEEKNUM', d(2021, 1, 1)],
    [2, 'WEEKNUM', d(2021, 1, 3)],
    [1, 'WEEKNUM', d(2021, 1, 3), 2],
    [2, 'WEEKNUM', d(2021, 1, 2), 16],
    [53, 'WEEKNUM', d(2021, 1, 3), 21],
    [10, 'WEEKNUM', d(2012, 3, 9), 21],
    ['#NUM!', 'WEEKNUM', d(2012, 3, 9), 3],
    [10, 'ISOWEEKNUM', d(2012, 3, 9)],
    [11, 'ISOWEEKNUM', 45366],
    [53, 'ISOWEEKNUM', d(2021, 1, 1)],
    [53, 'ISOWEEKNUM', d(2021, 1, 3)],
    [53, 'ISOWEEKNUM', d(2016, 1, 1)],
    [52, 'ISOWEEKNUM', d(2012, 12, 30)],
    [1, 'ISOWEEKNUM', d(2012, 12, 31)],
    [1, 'ISOWEEKNUM', d(2019, 12, 30)],
    [1, 'ISOWEEKNUM', d(2024, 12, 31)],
    [52, 'ISOWEEKNUM', d(2023, 12, 31)],
    [1, 'ISOWEEKNUM', d(2008, 12, 29)],
    [53, 'ISOWEEKNUM', d(2010, 1, 3)],
  ]);
  context = makeContext({}, true);
  // 1 January 1904 was a Friday.
  check([
    [6, 'WEEKDAY', 0],
    [6, 'WEEKDAY', 45366 - 1462],
  ]);
});

test('moving by months', () => {
  context = makeContext({});
  check([
    [d(2011, 2, 15), 'EDATE', d(2011, 1, 15), 1],
    [d(2010, 12, 15), 'EDATE', d(2011, 1, 15), -1],
    [d(2011, 3, 15), 'EDATE', d(2011, 1, 15), 2],
    [d(2024, 2, 29), 'EDATE', d(2024, 1, 31), 1],
    [d(2023, 2, 28), 'EDATE', d(2023, 1, 31), 1],
    [d(2025, 2, 28), 'EDATE', d(2024, 2, 29), 12],
    [d(2024, 3, 31), 'EDATE', d(2024, 1, 31), 2.9],
    ['#NUM!', 'EDATE', d(2024, 1, 31), -2000 * 12],
    ['#NUM!', 'EDATE', d(2024, 1, 31), 1e6 * 12],
    [d(2011, 2, 28), 'EOMONTH', d(2011, 1, 1), 1],
    [d(2010, 10, 31), 'EOMONTH', d(2011, 1, 1), -3],
    [d(2024, 2, 29), 'EOMONTH', d(2024, 2, 10), 0],
    [d(2024, 12, 31), 'EOMONTH', d(2024, 2, 10), 10],
    // February 1900 has 29 days in the 1900 system.
    [60, 'EOMONTH', 40, 0],
  ]);
});

test('differences between dates', () => {
  context = makeContext({});
  check([
    [2, 'DATEDIF', d(2001, 1, 1), d(2003, 1, 1), 'Y'],
    [440, 'DATEDIF', d(2001, 6, 1), d(2002, 8, 15), 'D'],
    [75, 'DATEDIF', d(2001, 6, 1), d(2002, 8, 15), 'YD'],
    [14, 'DATEDIF', d(2001, 6, 1), d(2002, 8, 15), 'M'],
    [2, 'DATEDIF', d(2001, 6, 1), d(2002, 8, 15), 'YM'],
    [14, 'DATEDIF', d(2001, 6, 1), d(2002, 8, 15), 'MD'],
    [1, 'DATEDIF', d(2001, 6, 1), d(2002, 8, 15), 'y'],
    [0, 'DATEDIF', d(2023, 3, 16), d(2024, 3, 15), 'Y'],
    [11, 'DATEDIF', d(2023, 3, 16), d(2024, 3, 15), 'M'],
    [11, 'DATEDIF', d(2023, 3, 16), d(2024, 3, 15), 'YM'],
    [1, 'DATEDIF', d(2023, 3, 15), d(2024, 3, 15), 'Y'],
    [0, 'DATEDIF', d(2023, 3, 15), d(2024, 3, 15), 'YD'],
    [0, 'DATEDIF', 45366, 45366, 'D'],
    ['#NUM!', 'DATEDIF', d(2024, 3, 15), d(2023, 3, 15), 'D'],
    ['#NUM!', 'DATEDIF', d(2023, 3, 15), d(2024, 3, 15), 'W'],
    [42, 'DAYS', '15-MAR-2021', '1-FEB-2021'],
    [364, 'DAYS', d(2011, 12, 31), d(2011, 1, 1)],
    [-364, 'DAYS', d(2011, 1, 1), d(2011, 12, 31)],
    ['#VALUE!', 'DAYS', 'x', 1],
    [1, 'DAYS360', d(2011, 1, 30), d(2011, 2, 1)],
    [360, 'DAYS360', d(2011, 1, 1), d(2011, 12, 31)],
    [30, 'DAYS360', d(2011, 1, 1), d(2011, 2, 1)],
    [60, 'DAYS360', d(2024, 1, 31), d(2024, 3, 31)],
    [57, 'DAYS360', d(2011, 1, 1), d(2011, 2, 28)],
    [-30, 'DAYS360', d(2011, 2, 1), d(2011, 1, 1)],
    [359, 'DAYS360', d(2011, 1, 1), d(2011, 12, 31), true],
    [60, 'DAYS360', d(2024, 1, 31), d(2024, 3, 31), true],
  ]);
  // UNSURE: DATEDIF "MD" when the end day is before the start day. Excel documents this unit as giving
  // negative or inaccurate results; this engine counts from the start day in the month before the end, which
  // can overshoot (31 January to 1 March: "31 February" is 2 March, one day after the end).
  check([
    [-1, 'DATEDIF', d(2024, 1, 31), d(2024, 3, 1), 'MD'],
    [16, 'DATEDIF', d(2024, 1, 28), d(2024, 3, 15), 'MD'],
    [24, 'DATEDIF', d(2024, 1, 20), d(2024, 3, 15), 'MD'],
  ]);
  // UNSURE: DAYS360 (US method) from the last day of February: that day counts as a 30th.
  check([
    [1, 'DAYS360', d(2011, 2, 28), d(2011, 3, 1)],
    [30, 'DAYS360', d(2011, 2, 28), d(2011, 3, 31)],
  ]);
  // UNSURE: DAYS with times of day: both sides are cut to whole days first.
  check([[14, 'DAYS', d(2021, 3, 15) + 0.5, d(2021, 3, 1)]]);
});

test('YEARFRAC on each basis', () => {
  context = makeContext({});
  const a = d(2012, 1, 1);
  const b = d(2012, 7, 30);
  check([
    [209 / 360, 'YEARFRAC', a, b],
    [209 / 360, 'YEARFRAC', a, b, 0],
    [211 / 366, 'YEARFRAC', a, b, 1],
    [211 / 360, 'YEARFRAC', a, b, 2],
    [211 / 365, 'YEARFRAC', a, b, 3],
    [209 / 360, 'YEARFRAC', a, b, 4],
    // The order of the dates does not matter.
    [209 / 360, 'YEARFRAC', b, a],
    [1, 'YEARFRAC', d(2023, 1, 1), d(2024, 1, 1), 1],
    [439 / 365.5, 'YEARFRAC', d(2023, 1, 1), d(2024, 3, 15), 1],
    [0, 'YEARFRAC', a, a, 1],
    ['#NUM!', 'YEARFRAC', a, b, 5],
    ['#NUM!', 'YEARFRAC', a, b, -1],
  ]);
});

test('working days', () => {
  context = makeContext({ H1: d(2012, 11, 22), H2: d(2012, 12, 4), H3: d(2013, 1, 21), H4: 'note', K1: d(2008, 11, 26), K2: d(2008, 12, 4), K3: d(2009, 1, 21) });
  check([
    [110, 'NETWORKDAYS', d(2012, 10, 1), d(2013, 3, 1)],
    [109, 'NETWORKDAYS', d(2012, 10, 1), d(2013, 3, 1), d(2012, 11, 22)],
    [107, 'NETWORKDAYS', d(2012, 10, 1), d(2013, 3, 1), ref('H1:H4')],
    [107, 'NETWORKDAYS', d(2012, 10, 1), d(2013, 3, 1), arr([[d(2012, 11, 22), d(2012, 12, 4), d(2013, 1, 21), d(2013, 1, 21)]])],
    [-110, 'NETWORKDAYS', d(2013, 3, 1), d(2012, 10, 1)],
    [23, 'NETWORKDAYS', d(2024, 1, 1), d(2024, 1, 31)],
    [1, 'NETWORKDAYS', 45366, 45366],
    [0, 'NETWORKDAYS', d(2024, 3, 16), d(2024, 3, 17)],
    // A holiday on a weekend takes nothing away.
    [1, 'NETWORKDAYS', 45366, d(2024, 3, 17), d(2024, 3, 16)],
    [5, 'NETWORKDAYS', d(2024, 3, 11), 45366],
    [6, 'NETWORKDAYS', d(2024, 3, 8), 45366],
    [d(2009, 4, 30), 'WORKDAY', d(2008, 10, 1), 151],
    [d(2009, 5, 5), 'WORKDAY', d(2008, 10, 1), 151, ref('K1:K3')],
    [d(2024, 3, 22), 'WORKDAY', 45366, 5],
    [d(2024, 3, 18), 'WORKDAY', 45366, 1],
    [d(2024, 3, 14), 'WORKDAY', 45366, -1],
    [d(2024, 3, 8), 'WORKDAY', 45366, -5],
    [d(2024, 3, 11), 'WORKDAY', d(2024, 3, 18), -5],
    [45366, 'WORKDAY', 45366, 0],
    // From a weekend the first working day is Monday (forward) or Friday (backward).
    [d(2024, 3, 18), 'WORKDAY', d(2024, 3, 16), 1],
    [d(2024, 3, 18), 'WORKDAY', d(2024, 3, 17), 1],
    [45366, 'WORKDAY', d(2024, 3, 17), -1],
    [45366, 'WORKDAY', d(2024, 3, 16), -1],
    // Holidays passed on the way push the result further, also when that reaches more holidays.
    [d(2024, 3, 20), 'WORKDAY', 45366, 1, arr([[d(2024, 3, 18), d(2024, 3, 19)]])],
    [d(2024, 3, 12), 'WORKDAY', 45366, -1, arr([[d(2024, 3, 14), d(2024, 3, 13)]])],
    ['#NUM!', 'WORKDAY', 45366, -1e6],
    ['#NUM!', 'WORKDAY', 45366, 1e300],
    ['#VALUE!', 'WORKDAY', 45366, 1, 'x'],
  ]);
  // Agreement with a day-by-day count over many spans.
  const date1904 = false;
  for (let start = 45290; start < 45320; start++) {
    for (let span = 0; span < 40; span++) {
      let expected = 0;
      for (let day = start; day <= start + span; day++) {
        const weekday = (day + 6) % 7;
        if (weekday !== 0 && weekday !== 6) expected++;
      }
      assert.equal(call('NETWORKDAYS', start, start + span), expected, `${start} + ${span} (${date1904})`);
      if (span > 0 && expected > 0) {
        const end = call('WORKDAY', start, span) as number;
        // WORKDAY and NETWORKDAYS agree: the days after the start up to the result are `span` working days.
        assert.equal(call('NETWORKDAYS', start + 1, end), span);
      }
    }
  }
});

test('DATEVALUE and TIMEVALUE', () => {
  context = makeContext({});
  check([
    [40777, 'DATEVALUE', '8/22/2011'],
    [40685, 'DATEVALUE', '22-MAY-2011'],
    [40597, 'DATEVALUE', '2011/02/23'],
    [45366, 'DATEVALUE', '2024-03-15'],
    [45366, 'DATEVALUE', '3/15/2024 6:35 AM'],
    // Without a year the date is in the year of the calculation's clock.
    [d(2024, 7, 5), 'DATEVALUE', '5-JUL'],
    [0, 'DATEVALUE', '6:35 AM'],
    ['#VALUE!', 'DATEVALUE', 'tomorrow'],
    ['#VALUE!', 'DATEVALUE', '45366'],
    ['#VALUE!', 'DATEVALUE', 45366],
    ['#VALUE!', 'DATEVALUE', ''],
    ['#N/A', 'DATEVALUE', err('#N/A')],
    [0.1, 'TIMEVALUE', '2:24 AM'],
    [hms(6, 35, 0), 'TIMEVALUE', '22-Aug-2011 6:35 AM'],
    [hms(18, 30, 0), 'TIMEVALUE', '6:30 PM'],
    [hms(16, 48, 10), 'TIMEVALUE', '16:48:10'],
    [0, 'TIMEVALUE', '8/22/2011'],
    ['#VALUE!', 'TIMEVALUE', 'noon'],
    ['#VALUE!', 'TIMEVALUE', 0.5],
  ]);
  // UNSURE: text with more than 24 hours. TIMEVALUE keeps only the fraction of a day.
  check([[hms(1, 0, 0), 'TIMEVALUE', '25:00']]);
});
