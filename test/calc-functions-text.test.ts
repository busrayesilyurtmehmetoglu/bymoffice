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

test('joining text', () => {
  context = makeContext({ A1: 'a', A2: 'b', A3: '', A5: 'c', B1: 1.5, B2: true, B3: err('#N/A') });
  check([
    ['ab', 'CONCATENATE', 'a', 'b'],
    ['a1.5TRUE', 'CONCATENATE', 'a', 1.5, true],
    ['x', 'CONCATENATE', 'x', null],
    ['#N/A', 'CONCATENATE', 'a', err('#N/A')],
    // CONCATENATE takes values one by one: a range gives one result per cell.
    [[['a!'], ['b!']], 'CONCATENATE', ref('A1:A2'), '!'],
    ['abc', 'CONCAT', ref('A1:A5')],
    ['ab1.5TRUE', 'CONCAT', ref('A1:A2'), ref('B1:B2')],
    ['12x', 'CONCAT', arr([[1, 2]]), 'x'],
    ['#N/A', 'CONCAT', ref('B1:B3')],
    ['a,b,c', 'TEXTJOIN', ',', true, ref('A1:A5')],
    ['a,b,,c', 'TEXTJOIN', ',', false, ref('A1:A3'), ref('A5')],
    ['a-b', 'TEXTJOIN', '-', true, 'a', 'b'],
    ['abc', 'TEXTJOIN', '', true, 'a', 'b', 'c'],
    ['a-b+c-d', 'TEXTJOIN', arr([['-', '+']]), true, 'a', 'b', 'c', 'd'],
    ['#N/A', 'TEXTJOIN', ',', true, ref('B1:B3')],
    ['', 'TEXTJOIN', ',', true, ref('C1:C3')],
  ]);
  const long = 'x'.repeat(20000);
  check([
    ['#VALUE!', 'CONCATENATE', long, long],
    ['#VALUE!', 'CONCAT', long, long],
    ['#VALUE!', 'TEXTJOIN', ',', true, long, long],
  ]);
});

test('cutting text', () => {
  check([
    ['S', 'LEFT', 'Sale Price'],
    ['Sale', 'LEFT', 'Sale Price', 4],
    ['Sale Price', 'LEFT', 'Sale Price', 100],
    ['', 'LEFT', 'Sale Price', 0],
    ['#VALUE!', 'LEFT', 'Sale Price', -1],
    ['12', 'LEFT', 1234, 2],
    ['TR', 'LEFT', true, 2],
    ['e', 'RIGHT', 'Sale Price'],
    ['Price', 'RIGHT', 'Sale Price', 5],
    ['Sale Price', 'RIGHT', 'Sale Price', 99],
    ['#VALUE!', 'RIGHT', 'x', -1],
    ['Fluid', 'MID', 'Fluid Flow', 1, 5],
    ['Flow', 'MID', 'Fluid Flow', 7, 20],
    ['', 'MID', 'Fluid Flow', 20, 5],
    ['', 'MID', 'Fluid Flow', 3, 0],
    ['#VALUE!', 'MID', 'Fluid Flow', 0, 5],
    ['#VALUE!', 'MID', 'Fluid Flow', 1, -1],
    [[['a', 'b', 'c']], 'MID', 'abc', arr([[1, 2, 3]]), 1],
    [11, 'LEN', 'Phoenix, AZ'],
    [0, 'LEN', ''],
    [0, 'LEN', null],
    [4, 'LEN', true],
    [3, 'LEN', 1.5],
    [[[1], [2]], 'LEN', arr([['a'], ['bb']])],
    ['#N/A', 'LEN', err('#N/A')],
  ]);
  context = makeContext({ A1: 'one', A2: 'three', A3: 12345 });
  check([[[[3], [5], [5]], 'LEN', ref('A1:A3')]]);
});

test('case and cleaning', () => {
  check([
    ['e. e. cummings', 'LOWER', 'E. E. Cummings'],
    ['TOTAL', 'UPPER', 'total'],
    ['This Is A Title', 'PROPER', 'this is a TITLE'],
    ['2-Way Street', 'PROPER', '2-way street'],
    ['76Budget', 'PROPER', '76BudGet'],
    ["It'S", 'PROPER', "it's"],
    ['First Quarter Earnings', 'TRIM', '  First   Quarter Earnings  '],
    ['', 'TRIM', '    '],
    // Only the space character is trimmed.
    ['\ta ', 'TRIM', ' \ta  '],
    ['text', 'CLEAN', '\u0007te\nxt\u0001'],
    [true, 'EXACT', 'word', 'word'],
    [false, 'EXACT', 'Word', 'word'],
    [true, 'EXACT', 1, '1'],
    ['#N/A', 'EXACT', err('#N/A'), 'a'],
  ]);
  // Language-neutral case mapping: the Turkish dotted and dotless i are not paired as Turkish expects.
  check([
    ['I', 'UPPER', 'i'],
    ['i', 'LOWER', 'I'],
  ]);
});

test('finding and replacing', () => {
  check([
    [1, 'FIND', 'M', 'Miriam McGovern'],
    [6, 'FIND', 'm', 'Miriam McGovern'],
    [8, 'FIND', 'M', 'Miriam McGovern', 3],
    ['#VALUE!', 'FIND', 'z', 'Miriam McGovern'],
    ['#VALUE!', 'FIND', 'M', 'Miriam', 0],
    ['#VALUE!', 'FIND', 'M', 'Miriam', 10],
    [1, 'FIND', '', 'abc'],
    [3, 'FIND', '', 'abc', 3],
    // FIND has no wildcards.
    ['#VALUE!', 'FIND', 'a*c', 'abc'],
    [4, 'FIND', '*', 'abc*'],
    [7, 'SEARCH', 'e', 'Statements', 6],
    [8, 'SEARCH', 'margin', 'Profit Margin'],
    [1, 'SEARCH', 'PROFIT', 'Profit Margin'],
    ['#VALUE!', 'SEARCH', 'x', 'Profit Margin'],
    [3, 'SEARCH', 'n?n', 'Banana'],
    [2, 'SEARCH', 'a*a', 'Banana'],
    [1, 'SEARCH', '*', 'Banana'],
    [4, 'SEARCH', '~*', 'abc*'],
    [2, 'SEARCH', '~?', 'a?c'],
    ['#VALUE!', 'SEARCH', 'a?a?a?a', 'Banana'],
    ['#VALUE!', 'SEARCH', 'a', 'Banana', 0],
    [4, 'SEARCH', 'a', 'Banana', 3],
    ['Sales Cost', 'SUBSTITUTE', 'Sales Data', 'Data', 'Cost'],
    ['Quarter 2, 2008', 'SUBSTITUTE', 'Quarter 1, 2008', '1', '2', 1],
    ['Quarter 1, 2012', 'SUBSTITUTE', 'Quarter 1, 2011', '1', '2', 3],
    ['bbb', 'SUBSTITUTE', 'aaa', 'a', 'b'],
    ['aaa', 'SUBSTITUTE', 'aaa', 'a', 'b', 4],
    ['aaa', 'SUBSTITUTE', 'aaa', 'A', 'b'],
    ['aaa', 'SUBSTITUTE', 'aaa', '', 'b'],
    ['', 'SUBSTITUTE', 'aaa', 'a', ''],
    ['#VALUE!', 'SUBSTITUTE', 'aaa', 'a', 'b', 0],
    ['abcde*k', 'REPLACE', 'abcdefghijk', 6, 5, '*'],
    ['2010', 'REPLACE', '2009', 3, 2, '10'],
    ['@456', 'REPLACE', '123456', 1, 3, '@'],
    ['abcX', 'REPLACE', 'abc', 10, 2, 'X'],
    ['Xabc', 'REPLACE', 'abc', 1, 0, 'X'],
    ['#VALUE!', 'REPLACE', 'abc', 0, 1, 'X'],
    ['*-*-*-', 'REPT', '*-', 3],
    ['', 'REPT', 'x', 0],
    ['xx', 'REPT', 'x', 2.9],
    ['#VALUE!', 'REPT', 'x', -1],
    ['#VALUE!', 'REPT', 'x', 32768],
    ['#VALUE!', 'REPT', 'ab', 16384],
    ['', 'REPT', '', 1e9],
  ]);
  assert.equal((call('REPT', 'x', 32767) as string).length, 32767);
  assert.equal(plain(call('SUBSTITUTE', 'a'.repeat(30000), 'a', 'bb')), '#VALUE!');
});

test('TEXTBEFORE and TEXTAFTER', () => {
  check([
    ['a', 'TEXTBEFORE', 'a-b-c', '-'],
    ['a-b', 'TEXTBEFORE', 'a-b-c', '-', 2],
    ['a-b', 'TEXTBEFORE', 'a-b-c', '-', -1],
    ['a', 'TEXTBEFORE', 'a-b-c', '-', -2],
    ['#N/A', 'TEXTBEFORE', 'a-b-c', '-', 3],
    ['#N/A', 'TEXTBEFORE', 'a-b-c', 'x'],
    ['none', 'TEXTBEFORE', 'a-b-c', 'x', 1, 0, 0, 'none'],
    ['#VALUE!', 'TEXTBEFORE', 'a-b-c', '-', 0],
    ['#VALUE!', 'TEXTBEFORE', 'a-b-c', '-', 9],
    ['a-b-c', 'TEXTBEFORE', 'a-b-c', '-', 3, 0, 1],
    ['#N/A', 'TEXTBEFORE', 'aXbxc', 'x', 2],
    ['aXb', 'TEXTBEFORE', 'aXbxc', 'x'],
    ['a', 'TEXTBEFORE', 'aXbxc', 'x', 1, 1],
    ['', 'TEXTBEFORE', 'abc', ''],
    ['b-c', 'TEXTAFTER', 'a-b-c', '-'],
    ['c', 'TEXTAFTER', 'a-b-c', '-', 2],
    ['c', 'TEXTAFTER', 'a-b-c', '-', -1],
    ['b-c', 'TEXTAFTER', 'a-b-c', '-', -2],
    ['#N/A', 'TEXTAFTER', 'a-b-c', '-', 3],
    ['', 'TEXTAFTER', 'a-b-c', '-', 3, 0, 1],
    ['c', 'TEXTAFTER', 'a::b::c', '::', 2],
    ['abc', 'TEXTAFTER', 'abc', ''],
    ['#N/A', 'TEXTAFTER', err('#N/A'), '-'],
  ]);
});

test('text to number', () => {
  context = makeContext({});
  check([
    [1000, 'VALUE', '$1,000'],
    [1234.5, 'VALUE', '1,234.5'],
    [0.12, 'VALUE', '12%'],
    [-5, 'VALUE', '(5)'],
    [1000, 'VALUE', '1e3'],
    [12, 'VALUE', ' 12 '],
    [7, 'VALUE', 7],
    [0.7, 'VALUE', '16:48:00'],
    [45366, 'VALUE', '3/15/2024'],
    [45366, 'VALUE', '2024-03-15'],
    ['#VALUE!', 'VALUE', 'abc'],
    ['#VALUE!', 'VALUE', ''],
    ['#VALUE!', 'VALUE', true],
    ['#VALUE!', 'VALUE', '=1+1'],
    [0, 'VALUE', null],
    ['#N/A', 'VALUE', err('#N/A')],
    [2500.27, 'NUMBERVALUE', '2.500,27', ',', '.'],
    [0.035, 'NUMBERVALUE', '3.5%'],
    [0.00035, 'NUMBERVALUE', '3.5%%'],
    [1234.5, 'NUMBERVALUE', '1,234.5'],
    [1234.5, 'NUMBERVALUE', ' 1 234,5 ', ',', '.'],
    [0, 'NUMBERVALUE', ''],
    [-12, 'NUMBERVALUE', '-12'],
    [1200, 'NUMBERVALUE', '1.2e3'],
    ['#VALUE!', 'NUMBERVALUE', '1.2.3'],
    ['#VALUE!', 'NUMBERVALUE', '1.5,000'],
    ['#VALUE!', 'NUMBERVALUE', 'abc'],
    ['#VALUE!', 'NUMBERVALUE', '5%5'],
    ['#VALUE!', 'NUMBERVALUE', '1,5', ',', ','],
    ['#VALUE!', 'NUMBERVALUE', '1', ''],
  ]);
});

test('number to text', () => {
  context = makeContext({});
  check([
    ['1,234.57', 'TEXT', 1234.567, '#,##0.00'],
    ['$1,234.57', 'TEXT', 1234.567, '$#,##0.00'],
    ['28.5%', 'TEXT', 0.285, '0.0%'],
    ['2024-03-15', 'TEXT', 45366, 'yyyy-mm-dd'],
    ['Friday', 'TEXT', 45366, 'dddd'],
    ['Mar-24', 'TEXT', 45366, 'mmm-yy'],
    ['12:00 PM', 'TEXT', 0.5, 'h:mm AM/PM'],
    ['0012', 'TEXT', 12, '0000'],
    ['1.23E+04', 'TEXT', 12345, '0.00E+00'],
    // Numeric text is formatted as its number, other text passes through the text section.
    ['12.00', 'TEXT', '12', '0.00'],
    ['abc', 'TEXT', 'abc', '0.00'],
    ['TRUE', 'TEXT', true, '0.00'],
    ['0.00', 'TEXT', null, '0.00'],
    ['5', 'TEXT', 5, 'General'],
    ['', 'TEXT', 5, ''],
    ['#N/A', 'TEXT', err('#N/A'), '0'],
    ['#N/A', 'TEXT', 1, err('#N/A')],
    ['#VALUE!', 'TEXT', -1, 'yyyy-mm-dd'],
    [[['1.0', '2.0']], 'TEXT', arr([[1, 2]]), '0.0'],
    ['1,234.6', 'FIXED', 1234.567, 1],
    ['1,230', 'FIXED', 1234.567, -1],
    ['-1230', 'FIXED', -1234.567, -1, true],
    ['44.33', 'FIXED', 44.332],
    ['1,234.57', 'FIXED', 1234.567],
    ['1234.567000', 'FIXED', 1234.567, 6, true],
    ['0.01', 'FIXED', 0.005],
    ['1.01', 'FIXED', 1.005, 2],
    ['1,000,000', 'FIXED', 999999.5, 0],
    ['#VALUE!', 'FIXED', 'abc'],
    ['$1,234.57', 'DOLLAR', 1234.567],
    ['$1,200', 'DOLLAR', 1234.567, -2],
    ['$0.12', 'DOLLAR', 0.123],
    ['$99.89', 'DOLLAR', 99.888],
  ]);
  // UNSURE: negative amounts. Excel's documentation shows parentheses, "($1,200)"; current en-US builds may show "-$1,200".
  check([
    ['($1,234.57)', 'DOLLAR', -1234.567, 2],
    ['($1,200)', 'DOLLAR', -1234.567, -2],
    ['($0.1230)', 'DOLLAR', -0.123, 4],
  ]);
  // Dates of the 1904 system.
  context = makeContext({}, true);
  check([['1904-01-02', 'TEXT', 1, 'yyyy-mm-dd']]);
});

test('characters and codes', () => {
  check([
    ['A', 'CHAR', 65],
    ['!', 'CHAR', 33],
    ['€', 'CHAR', 128],
    ['é', 'CHAR', 233],
    ['ÿ', 'CHAR', 255],
    ['#VALUE!', 'CHAR', 0],
    ['#VALUE!', 'CHAR', 256],
    [65, 'CODE', 'A'],
    [65, 'CODE', 'Alpha'],
    [33, 'CODE', '!'],
    [128, 'CODE', '€'],
    [233, 'CODE', 'é'],
    ['#VALUE!', 'CODE', ''],
    ['B', 'UNICHAR', 66],
    ['€', 'UNICHAR', 8364],
    ['😀', 'UNICHAR', 128512],
    ['#VALUE!', 'UNICHAR', 0],
    ['#VALUE!', 'UNICHAR', 1114112],
    [66, 'UNICODE', 'B'],
    [8364, 'UNICODE', '€'],
    [128512, 'UNICODE', '😀'],
    ['#VALUE!', 'UNICODE', ''],
  ]);
  // UNSURE: CODE of a character outside Windows-1252 answers 63 ("?"), and UNICHAR of a surrogate answers #N/A.
  check([
    [63, 'CODE', '中'],
    ['#N/A', 'UNICHAR', 0xd800],
  ]);
});
