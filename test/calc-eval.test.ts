import assert from 'node:assert/strict';
import { test } from 'node:test';
import { EvaluationAbort, evaluate, evaluationDepth, isBuiltinForm, refRange, toCellValue, type EvaluateOptions } from '../src/calc/eval';
import { parseFormula, type Node } from '../src/calc/parse';
import {
  calcError,
  isArray,
  isError,
  isRef,
  refToArray,
  toNumber,
  type ArrayValue,
  type CalcContext,
  type ErrorCode,
  type FunctionSpec,
  type Scalar,
  type Value,
} from '../src/calc/values';
import { parseCellRef, type CellRange } from '../src/model/sheet/types';

type Grid = Record<string, Scalar>;

/** A workbook that is nothing but values: sheet name -> "A1" -> value. */
function makeContext(sheets: Record<string, Grid>, at = 'Z100', sheet = 0): CalcContext {
  const names = Object.keys(sheets);
  const cells = names.map((name) => {
    const map = new Map<string, Scalar>();
    for (const address of Object.keys(sheets[name])) {
      const cell = parseCellRef(address);
      if (cell) map.set(`${cell.row}:${cell.col}`, sheets[name][address]);
    }
    return map;
  });
  const here = parseCellRef(at) as { row: number; col: number };
  return {
    sheet,
    row: here.row,
    col: here.col,
    date1904: false,
    cellValue: (s, row, col) => {
      const value = cells[s] ? cells[s].get(`${row}:${col}`) : undefined;
      return value === undefined ? null : value;
    },
    forEachCell: (s, range, visit) => {
      const found: [number, number, Scalar][] = [];
      if (cells[s]) {
        cells[s].forEach((value, key) => {
          const [row, col] = key.split(':').map(Number);
          if (row >= range.top && row <= range.bottom && col >= range.left && col <= range.right && value !== null) found.push([row, col, value]);
        });
      }
      found.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
      for (const [row, col, value] of found) visit(value, row, col);
    },
    usedExtent: (s) => {
      let rows = 0;
      let cols = 0;
      if (cells[s]) {
        cells[s].forEach((_value, key) => {
          const [row, col] = key.split(':').map(Number);
          rows = Math.max(rows, row + 1);
          cols = Math.max(cols, col + 1);
        });
      }
      return { rows, cols };
    },
    sheetIndex: (name) => {
      const index = names.findIndex((other) => other.toLowerCase() === name.toLowerCase());
      return index < 0 ? undefined : index;
    },
    sheetName: (index) => names[index],
    sheetCount: () => names.length,
    now: () => 45000.5,
    random: () => 0.25,
  };
}

/** Every scalar of the arguments, references and arrays opened up. */
function flatten(args: Value[], context: CalcContext): Scalar[] {
  const out: Scalar[] = [];
  for (const arg of args) {
    if (isRef(arg)) context.forEachCell(arg.sheet, arg.range, (value) => out.push(value));
    else if (isArray(arg)) out.push(...arg.data);
    else out.push(arg);
  }
  return out;
}

let calls = 0;

/** A small library; the real one is tested on its own. */
const FUNCTIONS = new Map<string, FunctionSpec>([
  [
    'SUM',
    {
      min: 1,
      call: (args, context) => {
        let total = 0;
        for (const value of flatten(args, context)) {
          if (isError(value)) return value;
          if (typeof value === 'number') total += value;
        }
        return total;
      },
    },
  ],
  ['COUNT', { min: 1, call: (args, context) => flatten(args, context).filter((value) => typeof value === 'number').length }],
  ['CONCAT', { min: 1, call: (args, context) => flatten(args, context).map((value) => (value === null ? '' : String(value))).join('') }],
  ['PAIR', { min: 2, max: 2, call: (args) => (isRef(args[1]) ? 'ref' : isArray(args[1]) ? 'array' : 'scalar') }],
  ['BOOM', { min: 0, call: () => { throw new Error('broken function'); } }],
  ['ABORT', { min: 0, call: () => { throw new EvaluationAbort(); } }],
  ['NAN', { min: 0, call: () => NaN }],
  ['NOTHING', { min: 0, call: () => undefined as unknown as Value }],
  ['TICK', { min: 0, max: 2, call: (args) => { calls++; return args.length > 0 ? args[0] : calls; } }],
  ['DEPTH', { min: 0, max: 0, call: () => evaluationDepth() }],
]);

const SHEETS: Record<string, Grid> = {
  Data: {
    A1: 1, A2: 2, A3: 3, A4: 'text', A5: true,
    B1: 10, B2: 20, B3: 30,
    C1: '12', C2: 'abc', C3: calcError('#N/A'), C4: '',
    D1: 'x', D2: false, D3: 0,
  },
  'Other Sheet': { A1: 100, B2: 'far' },
};

function value(formula: string, at = 'Z100', options?: EvaluateOptions, sheets = SHEETS): Value {
  return evaluate(parseFormula(formula), makeContext(sheets, at), FUNCTIONS, options);
}

/** What a cell holding the formula would show. */
function cell(formula: string, at = 'Z100', options?: EvaluateOptions, sheets = SHEETS): Scalar {
  const context = makeContext(sheets, at);
  return toCellValue(evaluate(parseFormula(formula), context, FUNCTIONS, options), context);
}

const err = (code: ErrorCode) => calcError(code);

/** The formula is an array formula: ranges are arrays everywhere. */
const ARRAY: EvaluateOptions = { arrayFormula: true };

function array(formula: string, at = 'Z100', options?: EvaluateOptions): Scalar[][] {
  const context = makeContext(SHEETS, at);
  let result = evaluate(parseFormula(formula), context, FUNCTIONS, options);
  if (isRef(result)) result = refToArray(result, context);
  assert.ok(isArray(result), `${formula} is not an array`);
  const out: Scalar[][] = [];
  const data = (result as ArrayValue).data;
  for (let r = 0; r < (result as ArrayValue).rows; r++) out.push(data.slice(r * (result as ArrayValue).cols, (r + 1) * (result as ArrayValue).cols));
  return out;
}

test('arithmetic', () => {
  assert.equal(cell('1+2*3'), 7);
  assert.equal(cell('(1+2)*3'), 9);
  assert.equal(cell('7-10'), -3);
  assert.equal(cell('1/4'), 0.25);
  assert.equal(cell('2^10'), 1024);
  assert.equal(cell('-2^2'), 4);
  assert.equal(cell('2^3^2'), 64);
  assert.equal(cell('2^-1'), 0.5);
  assert.equal(cell('50%'), 0.5);
  assert.equal(cell('200%%'), 0.02);
  assert.equal(cell('-A1'), -1);
  assert.equal(cell('--A2'), 2);
  assert.equal(cell('+A4'), 'text');
  assert.equal(cell('A1+A2*B3'), 61);
  assert.deepEqual(cell('1/0'), err('#DIV/0!'));
  assert.deepEqual(cell('0/0'), err('#DIV/0!'));
  assert.deepEqual(cell('0^0'), err('#NUM!'));
  assert.deepEqual(cell('0^-1'), err('#DIV/0!'));
  assert.deepEqual(cell('(-8)^0.5'), err('#NUM!'));
  assert.equal(cell('(-8)^2'), 64);
  assert.deepEqual(cell('1E308*10'), err('#NUM!'));
  assert.deepEqual(cell('10^400'), err('#NUM!'));
  assert.ok(Object.is(cell('-0'), 0));
});

test('conversions between kinds of values', () => {
  assert.equal(cell('"12"+1'), 13);
  assert.equal(cell('C1*2'), 24);
  assert.equal(cell('"5%"+0'), 0.05);
  assert.equal(cell('TRUE+TRUE'), 2);
  assert.equal(cell('A5+1'), 2);
  assert.equal(cell('Z99+1'), 1);
  assert.deepEqual(cell('"abc"+1'), err('#VALUE!'));
  assert.deepEqual(cell('C2*1'), err('#VALUE!'));
  assert.deepEqual(cell('-"x"'), err('#VALUE!'));
  assert.deepEqual(cell('C4+1'), err('#VALUE!'));
  assert.equal(cell('1&2'), '12');
  assert.equal(cell('"a"&TRUE&1.5'), 'aTRUE1.5');
  assert.equal(cell('A1&Z99&"|"'), '1|');
  assert.equal(cell('0.1+0.2&""'), '0.3');
  assert.deepEqual(cell('REPT&1'), err('#NAME?'));
});

test('comparisons', () => {
  assert.equal(cell('1=1'), true);
  assert.equal(cell('1<>1'), false);
  assert.equal(cell('1<2'), true);
  assert.equal(cell('2<=2'), true);
  assert.equal(cell('1>2'), false);
  assert.equal(cell('2>=3'), false);
  assert.equal(cell('"a"="A"'), true);
  assert.equal(cell('"a"<"b"'), true);
  // Numbers sort before text, text before booleans; no conversion is made.
  assert.equal(cell('"1"=1'), false);
  assert.equal(cell('1<"a"'), true);
  assert.equal(cell('"z"<TRUE'), true);
  assert.equal(cell('Z99=0'), true);
  assert.equal(cell('Z99=""'), true);
  assert.equal(cell('Z99=FALSE'), true);
  assert.equal(cell('1+1=2'), true);
  assert.equal(cell('"a"&"b"="ab"'), true);
});

test('errors travel through operators, the left one first', () => {
  assert.deepEqual(cell('#N/A+1'), err('#N/A'));
  assert.deepEqual(cell('1+#REF!'), err('#REF!'));
  assert.deepEqual(cell('#NUM!+#N/A'), err('#NUM!'));
  assert.deepEqual(cell('#N/A&"x"'), err('#N/A'));
  assert.deepEqual(cell('"x"&#DIV/0!'), err('#DIV/0!'));
  assert.deepEqual(cell('#N/A=#N/A'), err('#N/A'));
  assert.deepEqual(cell('1<#NULL!'), err('#NULL!'));
  assert.deepEqual(cell('-#NAME?'), err('#NAME?'));
  assert.deepEqual(cell('#VALUE!%'), err('#VALUE!'));
  assert.deepEqual(cell('C3*2'), err('#N/A'));
  assert.deepEqual(cell('"abc"+#N/A'), err('#VALUE!'));
});

test('arrays are worked element by element', () => {
  assert.deepEqual(array('{1,2;3,4}*2'), [[2, 4], [6, 8]]);
  assert.deepEqual(array('{1,2;3,4}+{10,20;30,40}'), [[11, 22], [33, 44]]);
  // A single row or column stands for every row or column.
  assert.deepEqual(array('{1;2}+{10,20}'), [[11, 21], [12, 22]]);
  assert.deepEqual(array('{1,2;3,4}+{10;20}'), [[11, 12], [23, 24]]);
  assert.deepEqual(array('{1,2,3}+{1,2}'), [[2, 4, err('#N/A')]]);
  assert.deepEqual(array('-{1,-2}'), [[-1, 2]]);
  assert.deepEqual(array('{50,25}%'), [[0.5, 0.25]]);
  assert.deepEqual(array('{1,2}&"x"'), [['1x', '2x']]);
  assert.deepEqual(array('{1,2,3}>=2'), [[false, true, true]]);
  assert.deepEqual(array('{1,"a",#N/A}+1'), [[2, err('#VALUE!'), err('#N/A')]]);
  assert.equal(cell('SUM({1,2;3,4})'), 10);
  assert.equal(cell('SUM({1,2;3,4}*2)'), 20);
  // In one cell an array shows its first element.
  assert.equal(cell('{5,6}*2'), 10);
});

test('a range is an array in an array formula and one cell in an ordinary formula', () => {
  // In an array formula a reference to several cells behaves as the array of its values.
  assert.deepEqual(array('A1:A3*B1:B3', 'Z100', ARRAY), [[10], [40], [90]]);
  assert.deepEqual(array('A1:B2+1', 'Z100', ARRAY), [[2, 11], [3, 21]]);
  assert.deepEqual(array('-A1:A3', 'Z100', ARRAY), [[-1], [-2], [-3]]);
  assert.equal(cell('SUM(A1:A3*B1:B3)', 'Z100', ARRAY), 140);
  assert.equal(cell('A1:A3*2', 'Z100', ARRAY), 2);
  // In an ordinary formula it is the cell in the formula's row or column (implicit intersection).
  assert.equal(value('A1:A3*B1:B3', 'F2'), 40);
  assert.equal(value('A1:A3*2', 'F3'), 6);
  assert.equal(value('-A1:A3', 'F3'), -3);
  assert.equal(value('A1:B1&"!"', 'B9'), '10!');
  assert.equal(cell('SUM(A1:A3*B1:B3)', 'F2'), 40);
  assert.deepEqual(value('A1:A3*2'), err('#VALUE!'));
  assert.deepEqual(value('A1:B2+1', 'A1'), err('#VALUE!'));
  assert.deepEqual(cell('SUM(A1:A3*B1:B3)'), err('#VALUE!'));
  // An array stays an array.
  assert.deepEqual(array('{1,2}*A1:A3', 'F2'), [[2, 4]]);
});

test('arguments that take arrays or single values', () => {
  const library = new Map(FUNCTIONS);
  const kind = (arg: Value) => (isRef(arg) ? 'ref' : isArray(arg) ? 'array' : isError(arg) ? arg.error : String(arg));
  library.set('KINDS', { min: 1, call: (args) => args.map(kind).join(',') });
  library.set('ARRAYS', { min: 1, arrayArgs: true, call: (args) => args.map(kind).join(',') });
  library.set('SECOND', { min: 1, arrayArgs: (index) => index === 1, call: (args) => args.map(kind).join(',') });
  library.set('SCALARS', { min: 1, scalarArgs: true, call: (args) => args.map(kind).join(',') });
  library.set('FIRSTONLY', { min: 1, scalarArgs: (index) => index === 0, call: (args) => args.map(kind).join(',') });
  const run = (formula: string, at = 'F2', options?: EvaluateOptions) => evaluate(parseFormula(formula), makeContext(SHEETS, at), library, options);
  // No statement about its arguments: a range arrives as it is, an expression is calculated the ordinary way.
  assert.equal(run('KINDS(A1:A3,A1:A3*2,{1,2})'), 'ref,4,array');
  // Array arguments: the expression is calculated element by element.
  assert.equal(run('ARRAYS(A1:A3,A1:A3*2,-A1:A3)'), 'ref,array,array');
  assert.equal(run('SECOND(A1:A3*2,A1:A3*2)'), '4,array');
  // ... all the way down.
  assert.equal(run('ARRAYS(KINDS(A1:A3*2))'), 'array');
  assert.equal(run('ARRAYS(IF(A1:A3>1,1,0))'), 'array');
  // Single values: a range is intersected before the function sees it.
  assert.equal(run('SCALARS(A1:A3,A1:A3*2,A1,5)'), '2,4,1,5');
  assert.equal(run('SCALARS(A1:A3)', 'F9'), '#VALUE!');
  assert.equal(run('FIRSTONLY(A1:A3,A1:A3)'), '2,ref');
  // In an array formula nothing is intersected.
  assert.equal(run('SCALARS(A1:A3,A1:A3*2)', 'F2', ARRAY), 'ref,array');
  assert.equal(run('KINDS(A1:A3*2)', 'F2', ARRAY), 'array');
  // The mode of the outer formula comes back after an array argument.
  assert.equal(run('SECOND(1,A1:A3*2)&A1:A3', 'F3'), '1,array3');
});

test('references and implicit intersection', () => {
  assert.ok(isRef(value('A1')));
  assert.deepEqual(value('A1:B2'), { ref: true, sheet: 0, range: { top: 0, left: 0, bottom: 1, right: 1 } });
  assert.equal(cell('A1'), 1);
  assert.equal(cell('Z99'), null);
  assert.equal(cell("'Other Sheet'!A1"), 100);
  assert.equal(cell("'other sheet'!A1+A1"), 101);
  assert.deepEqual(cell('Missing!A1'), err('#REF!'));
  assert.deepEqual(cell('[1]Data!A1'), err('#REF!'));
  assert.deepEqual(cell('Data:Other!A1'), err('#REF!'));
  assert.deepEqual(cell('SUM(Missing!A1:A3)'), err('#REF!'));
  // A column gives the cell in the formula's row, a row the cell in its column.
  assert.equal(cell('A1:A3', 'F2'), 2);
  assert.equal(cell('A:A', 'F3'), 3);
  assert.equal(cell('A1:C1', 'B9'), 10);
  assert.equal(cell('1:1', 'B9'), 10);
  assert.deepEqual(cell('A1:A3', 'F9'), err('#VALUE!'));
  assert.deepEqual(cell('A1:B3', 'A2'), err('#VALUE!'));
  // Functions get the reference itself, or an array, never a reduced value.
  assert.equal(cell('PAIR(1,A1:A3)'), 'ref');
  assert.equal(cell('PAIR(1,A1)'), 'ref');
  assert.equal(cell('PAIR(1,A1:A3*1)', 'Z100', ARRAY), 'array');
  assert.equal(cell('PAIR(1,A1:A3*1)', 'F2'), 'scalar');
  assert.equal(cell('PAIR(1,A1*1)'), 'scalar');
  assert.equal(cell('SUM(A:A)'), 6);
  assert.equal(cell('SUM(1:1)'), 11);
});

test('reference operators', () => {
  assert.deepEqual(value('A1:INDEX(B1:B3,2)'), { ref: true, sheet: 0, range: { top: 0, left: 0, bottom: 1, right: 1 } });
  assert.equal(cell('SUM(A1:INDEX(A1:A3,2))'), 3);
  assert.equal(cell('SUM(INDEX(A1:A3,2):A3)'), 5);
  assert.deepEqual(cell("SUM(A1:INDEX('Other Sheet'!A1:A3,2))"), err('#REF!'));
  assert.deepEqual(cell('SUM(A1:5)'), err('#VALUE!'));
  // A space is the intersection.
  assert.deepEqual(value('A1:B2 B2:C3'), { ref: true, sheet: 0, range: { top: 1, left: 1, bottom: 1, right: 1 } });
  assert.equal(cell('A1:B2 B2:C3'), 20);
  assert.equal(cell('SUM(A:A 2:3)'), 5);
  assert.deepEqual(cell('A1:A2 B1:B2'), err('#NULL!'));
  assert.deepEqual(cell("A1:A2 'Other Sheet'!A1:A2"), err('#NULL!'));
  // A union is taken apart by functions that accept any number of arguments.
  assert.equal(cell('SUM((A1,B1:B2))'), 31);
  assert.equal(cell('SUM((A1,A2,A3),1)'), 7);
  assert.equal(cell('COUNT((A1:A3,B1:B3))'), 6);
  assert.deepEqual(cell('(A1,B1)'), err('#VALUE!'));
  // A function with a fixed number of arguments gets the union as the error it is.
  assert.deepEqual(cell('TICK((A1,B1))'), err('#VALUE!'));
  assert.equal(cell('AREAS((A1,B1:B2,C1))'), 3);
  assert.equal(cell('AREAS(A1:B2)'), 1);
  assert.deepEqual(cell('AREAS(5)'), err('#VALUE!'));
});

test('names', () => {
  const names: Record<string, string> = {
    RATE: '0.5',
    TOTAL: 'Data!$A$1:$A$3',
    LOCAL: 'A1',
    LOOP: 'LOOP+1',
    PING: 'PONG+1',
    PONG: 'PING+1',
    DOUBLE: 'RATE*2',
  };
  const seen: [string, number][] = [];
  const options: EvaluateOptions = {
    resolveName: (name, sheet) => {
      seen.push([name, sheet]);
      const text = names[name.toUpperCase()];
      return text === undefined ? undefined : parseFormula(text);
    },
  };
  assert.equal(cell('Rate*4', 'Z100', options), 2);
  assert.equal(cell('SUM(Total)', 'Z100', options), 6);
  assert.equal(cell('Double+1', 'Z100', options), 2);
  assert.equal(cell('Local', 'Z100', options), 1);
  assert.deepEqual(cell('Unknown+1', 'Z100', options), err('#NAME?'));
  assert.deepEqual(cell('Unknown+1'), err('#NAME?'));
  // A name that leads back to itself ends, with an error.
  assert.deepEqual(cell('Loop', 'Z100', options), err('#NAME?'));
  assert.deepEqual(cell('Ping', 'Z100', options), err('#NAME?'));
  seen.length = 0;
  assert.equal(cell("'Other Sheet'!Rate", 'Z100', options), 0.5);
  assert.deepEqual(seen, [['Rate', 1]]);
  assert.deepEqual(cell('Missing!Rate', 'Z100', options), err('#REF!'));
});

test('function calls', () => {
  assert.equal(cell('SUM(1,2,3)'), 6);
  assert.equal(cell('sum(A1:A3,B1)'), 16);
  assert.equal(cell('_xlfn.SUM(1)'), 1);
  assert.deepEqual(cell('NOSUCH(1)'), err('#NAME?'));
  // Too few or too many arguments.
  assert.deepEqual(cell('SUM()'), err('#VALUE!'));
  assert.deepEqual(cell('PAIR(1)'), err('#VALUE!'));
  assert.deepEqual(cell('PAIR(1,2,3)'), err('#VALUE!'));
  // An empty argument arrives as null.
  assert.equal(cell('CONCAT("a",,"b")'), 'ab');
  assert.equal(cell('TICK(,)'), null);
  // Whatever a function does wrong becomes an error value.
  assert.deepEqual(cell('BOOM()'), err('#VALUE!'));
  assert.deepEqual(cell('1+BOOM()'), err('#VALUE!'));
  assert.deepEqual(cell('NAN()'), err('#NUM!'));
  assert.deepEqual(cell('NOTHING()'), err('#VALUE!'));
  assert.equal(cell('IFERROR(BOOM(),"caught")'), 'caught');
  // Errors are handed to functions, not stopped before them.
  assert.deepEqual(cell('SUM(1,#N/A)'), err('#N/A'));
  // The one exception that is let through.
  assert.throws(() => cell('1+SUM(ABORT())'), EvaluationAbort);
  assert.equal(evaluationDepth(), 0);
});

test('IF and laziness', () => {
  assert.equal(cell('IF(TRUE,1,2)'), 1);
  assert.equal(cell('IF(FALSE,1,2)'), 2);
  assert.equal(cell('IF(A1>0,"pos","neg")'), 'pos');
  assert.equal(cell('IF(0,1)'), false);
  assert.equal(cell('IF(1,,2)'), 0);
  assert.equal(cell('IF(0,1,)'), 0);
  assert.equal(cell('IF("true",1,2)'), 1);
  assert.equal(cell('IF(Z99,1,2)'), 2);
  assert.deepEqual(cell('IF("abc",1,2)'), err('#VALUE!'));
  assert.deepEqual(cell('IF(#N/A,1,2)'), err('#N/A'));
  assert.deepEqual(cell('IF(TRUE)'), err('#VALUE!'));
  assert.deepEqual(cell('IF(1,2,3,4)'), err('#VALUE!'));
  // The branch not taken is not calculated: no error, no call.
  assert.equal(cell('IF(TRUE,1,1/0)'), 1);
  assert.equal(cell('IF(FALSE,BOOM(),"ok")'), 'ok');
  calls = 0;
  assert.equal(cell('IF(A1=1,"yes",TICK())'), 'yes');
  assert.equal(calls, 0);
  assert.equal(cell('IF(A1=2,"yes",TICK())'), 1);
  assert.equal(calls, 1);
  // IF hands on a reference as a reference.
  assert.equal(cell('SUM(IF(TRUE,A1:A3,B1:B3))'), 6);
  assert.equal(cell('PAIR(1,IF(FALSE,A1:A3,B1:B3))'), 'ref');
  // An array of conditions chooses element by element.
  assert.deepEqual(array('IF({1,0,1},"y","n")'), [['y', 'n', 'y']]);
  assert.deepEqual(array('IF(A1:A3>1,A1:A3,0)', 'Z100', ARRAY), [[0], [2], [3]]);
  assert.deepEqual(array('IF({0;1;1},A1:A3,0)'), [[0], [2], [3]]);
  // In an ordinary formula the condition is one cell, and the branch comes back as the reference it is.
  assert.equal(cell('IF(A1:A3>1,"big","small")', 'F1'), 'small');
  assert.equal(cell('IF(A1:A3>1,"big","small")', 'F3'), 'big');
  assert.equal(cell('SUM(IF(A1:A3>1,B1:B3))', 'F3'), 60);
  assert.deepEqual(array('IF({TRUE,FALSE},{1,2},{10,20})'), [[1, 20]]);
  assert.deepEqual(array('IF({TRUE,FALSE},FALSE,TRUE)'), [[false, true]]);
  assert.deepEqual(array('IF({1,"x",0},1)'), [[1, err('#VALUE!'), false]]);
  assert.equal(cell('SUM(IF(A1:A3>1,B1:B3))', 'Z100', ARRAY), 50);
  calls = 0;
  assert.deepEqual(array('IF({1,1},"y",TICK())'), [['y', 'y']]);
  assert.equal(calls, 0);
});

test('IFERROR, IFNA, IFS, SWITCH and CHOOSE', () => {
  assert.equal(cell('IFERROR(1/0,"div")'), 'div');
  assert.equal(cell('IFERROR(5,"div")'), 5);
  assert.equal(cell('IFERROR(C3,0)'), 0);
  assert.equal(cell('IFERROR(A1,0)'), 1);
  assert.deepEqual(array('IFERROR({1,#N/A,3}/1,-1)'), [[1, -1, 3]]);
  assert.deepEqual(cell('IFERROR(1)'), err('#VALUE!'));
  assert.equal(cell('IFNA(#N/A,"na")'), 'na');
  assert.deepEqual(cell('IFNA(1/0,"na")'), err('#DIV/0!'));
  assert.equal(cell('IFNA(7,"na")'), 7);
  calls = 0;
  assert.equal(cell('IFERROR(1,TICK())'), 1);
  assert.equal(cell('IFNA(1,TICK())'), 1);
  assert.equal(calls, 0);

  assert.equal(cell('IFS(FALSE,1,TRUE,2,TRUE,3)'), 2);
  assert.equal(cell('IFS(A1>5,"big",A1>0,"small")'), 'small');
  assert.deepEqual(cell('IFS(FALSE,1)'), err('#N/A'));
  assert.deepEqual(cell('IFS(TRUE)'), err('#VALUE!'));
  assert.deepEqual(cell('IFS(#REF!,1,TRUE,2)'), err('#REF!'));
  calls = 0;
  assert.equal(cell('IFS(TRUE,1,TICK(),2)'), 1);
  assert.equal(cell('IFS(FALSE,TICK(),TRUE,2)'), 2);
  assert.equal(calls, 0);

  assert.equal(cell('SWITCH(2,1,"one",2,"two",3,"three")'), 'two');
  assert.equal(cell('SWITCH("B","a","x","b","y")'), 'y');
  assert.equal(cell('SWITCH(9,1,"one","other")'), 'other');
  assert.deepEqual(cell('SWITCH(9,1,"one")'), err('#N/A'));
  assert.deepEqual(cell('SWITCH(#N/A,1,"one")'), err('#N/A'));
  assert.deepEqual(cell('SWITCH(1,2)'), err('#VALUE!'));
  calls = 0;
  assert.equal(cell('SWITCH(1,1,"one",2,TICK(),TICK())'), 'one');
  assert.equal(calls, 0);

  assert.equal(cell('CHOOSE(2,"a","b","c")'), 'b');
  assert.equal(cell('CHOOSE(2.9,"a","b","c")'), 'b');
  assert.equal(cell('SUM(CHOOSE(2,A1:A3,B1:B3))'), 60);
  assert.deepEqual(cell('CHOOSE(0,"a")'), err('#VALUE!'));
  assert.deepEqual(cell('CHOOSE(4,"a","b","c")'), err('#VALUE!'));
  assert.deepEqual(cell('CHOOSE("x","a")'), err('#VALUE!'));
  calls = 0;
  assert.equal(cell('CHOOSE(1,"a",TICK(),1/0)'), 'a');
  assert.equal(calls, 0);
});

test('AND, OR, XOR and NOT', () => {
  assert.equal(cell('AND(TRUE,1,"true")'), true);
  assert.equal(cell('AND(TRUE,0)'), false);
  assert.equal(cell('OR(FALSE,0)'), false);
  assert.equal(cell('OR(FALSE,2)'), true);
  assert.equal(cell('XOR(TRUE,TRUE,TRUE)'), true);
  assert.equal(cell('XOR(TRUE,TRUE)'), false);
  assert.equal(cell('AND(A1:A3)'), true);
  assert.equal(cell('AND(A1:A3,D3)'), false);
  // Inside a reference or an array, text and empty cells are passed over.
  assert.equal(cell('AND(A1:A5)'), true);
  assert.equal(cell('OR(D1:D3)'), false);
  assert.equal(cell('AND({1,"x",TRUE})'), true);
  assert.deepEqual(cell('AND(D1)'), err('#VALUE!'));
  assert.deepEqual(cell('OR(Y1:Y9)'), err('#VALUE!'));
  // Text given directly is read, and refused when it is not a boolean.
  assert.equal(cell('OR("FALSE","true")'), true);
  assert.deepEqual(cell('AND("abc",TRUE)'), err('#VALUE!'));
  assert.equal(cell('AND(TRUE,)'), false);
  assert.deepEqual(cell('AND(TRUE,#N/A)'), err('#N/A'));
  assert.deepEqual(cell('OR(TRUE,C1:C3)'), err('#N/A'));
  assert.deepEqual(cell('AND()'), err('#VALUE!'));
  assert.equal(cell('NOT(TRUE)'), false);
  assert.equal(cell('NOT(0)'), true);
  assert.equal(cell('NOT(Z99)'), true);
  assert.deepEqual(cell('NOT("x")'), err('#VALUE!'));
  assert.deepEqual(cell('NOT(1,2)'), err('#VALUE!'));
  assert.deepEqual(array('NOT({TRUE,FALSE})'), [[false, true]]);
});

test('ROW, COLUMN, ROWS, COLUMNS', () => {
  assert.equal(cell('ROW()', 'C7'), 7);
  assert.equal(cell('COLUMN()', 'C7'), 3);
  assert.equal(cell('ROW(B5)'), 5);
  assert.equal(cell('COLUMN(B5)'), 2);
  assert.equal(cell("ROW('Other Sheet'!A9)"), 9);
  // Every row of a range where an array is wanted: in an array formula, in an array argument.
  assert.deepEqual(array('ROW(A2:B4)', 'Z100', ARRAY), [[2], [3], [4]]);
  assert.deepEqual(array('COLUMN(B1:D1)', 'Z100', ARRAY), [[2, 3, 4]]);
  assert.equal(cell('SUM(ROW(A1:A4))', 'Z100', ARRAY), 10);
  // UNSURE: an ordinary formula gets the first row only (Excel 2019 and before: =SUM(ROW(1:10)) is 1 unless array-entered).
  assert.equal(value('ROW(A2:B4)'), 2);
  assert.equal(value('COLUMN(B1:D1)'), 2);
  assert.equal(cell('SUM(ROW(A1:A4))'), 1);
  assert.deepEqual(cell('ROW(5)'), err('#VALUE!'));
  assert.deepEqual(cell('ROW(Missing!A1)'), err('#REF!'));
  assert.equal(cell('ROWS(A1:C5)'), 5);
  assert.equal(cell('COLUMNS(A1:C5)'), 3);
  assert.equal(cell('ROWS(A:A)'), 1048576);
  assert.equal(cell('COLUMNS(1:1)'), 16384);
  assert.equal(cell('ROWS({1,2;3,4;5,6})'), 3);
  assert.equal(cell('COLUMNS({1,2;3,4;5,6})'), 2);
  assert.equal(cell('ROWS(5)'), 1);
  assert.deepEqual(cell('ROWS()'), err('#VALUE!'));
});

test('OFFSET', () => {
  assert.deepEqual(value('OFFSET(A1,1,1)'), { ref: true, sheet: 0, range: { top: 1, left: 1, bottom: 1, right: 1 } });
  assert.equal(cell('OFFSET(A1,1,1)'), 20);
  assert.equal(cell('SUM(OFFSET(A1,0,0,3,1))'), 6);
  assert.equal(cell('SUM(OFFSET(A1:A2,1,1))'), 50);
  assert.equal(cell('SUM(OFFSET(A1,0,0,3,2))'), 66);
  assert.equal(cell('SUM(OFFSET(A3,0,0,-2,1))'), 5);
  assert.equal(cell('SUM(OFFSET(A1,1,,2))'), 5);
  assert.equal(cell('OFFSET(A1,"1",1.9)'), 20);
  assert.equal(cell("OFFSET('Other Sheet'!A1,1,1)"), 'far');
  assert.deepEqual(cell('OFFSET(A1,-1,0)'), err('#REF!'));
  assert.deepEqual(cell('OFFSET(A1,0,-1)'), err('#REF!'));
  assert.deepEqual(cell('OFFSET(A1,1048576,0)'), err('#REF!'));
  assert.deepEqual(cell('OFFSET(A1,0,0,0,1)'), err('#REF!'));
  assert.deepEqual(cell('OFFSET(5,0,0)'), err('#VALUE!'));
  assert.deepEqual(cell('OFFSET(A1,"x",0)'), err('#VALUE!'));
  assert.deepEqual(cell('OFFSET(A1,1)'), err('#VALUE!'));
});

test('INDIRECT', () => {
  assert.equal(cell('INDIRECT("A2")'), 2);
  assert.equal(cell('INDIRECT("B"&3)'), 30);
  assert.equal(cell('SUM(INDIRECT("A1:A3"))'), 6);
  assert.equal(cell('SUM(INDIRECT("A:A"))'), 6);
  assert.equal(cell('INDIRECT("\'Other Sheet\'!A1")'), 100);
  assert.equal(cell('INDIRECT("$B$2",TRUE)'), 20);
  assert.equal(cell('INDIRECT(D1&"1")', 'Z100', undefined, { S: { D1: 'B', B1: 77 } }), 77);
  assert.deepEqual(cell('INDIRECT("R1C1",FALSE)'), err('#REF!'));
  assert.deepEqual(cell('INDIRECT("1+1")'), err('#REF!'));
  assert.deepEqual(cell('INDIRECT("=A1")'), err('#REF!'));
  assert.deepEqual(cell('INDIRECT("SUM(A1)")'), err('#REF!'));
  assert.deepEqual(cell('INDIRECT("")'), err('#REF!'));
  assert.deepEqual(cell('INDIRECT("not a ref")'), err('#REF!'));
  assert.deepEqual(cell('INDIRECT("Missing!A1")'), err('#REF!'));
  assert.deepEqual(cell('INDIRECT(#N/A)'), err('#N/A'));
  assert.deepEqual(cell('INDIRECT()'), err('#VALUE!'));
  const options: EvaluateOptions = { resolveName: (name) => (name.toUpperCase() === 'TOTAL' ? parseFormula('Data!$B$1:$B$3') : undefined) };
  assert.equal(cell('SUM(INDIRECT("Total"))', 'Z100', options), 60);
  assert.deepEqual(cell('INDIRECT("Nothing")', 'Z100', options), err('#REF!'));
});

test('INDEX', () => {
  assert.equal(cell('INDEX(A1:B3,2,2)'), 20);
  assert.equal(cell('INDEX(A1:A3,3)'), 3);
  assert.equal(cell('INDEX(A1:C1,2)'), 10);
  assert.ok(isRef(value('INDEX(A1:B3,2,2)')));
  assert.deepEqual(value('INDEX(A1:B3,0,2)'), { ref: true, sheet: 0, range: { top: 0, left: 1, bottom: 2, right: 1 } });
  assert.deepEqual(value('INDEX(A1:B3,2,0)'), { ref: true, sheet: 0, range: { top: 1, left: 0, bottom: 1, right: 1 } });
  assert.deepEqual(value('INDEX(A1:B3,2)'), { ref: true, sheet: 0, range: { top: 1, left: 0, bottom: 1, right: 1 } });
  assert.equal(cell('SUM(INDEX(A1:B3,0,2))'), 60);
  assert.equal(cell('SUM(INDEX(A1:B3,2,))'), 22);
  assert.equal(cell('INDEX(A1:B3,2,2,1)'), 20);
  assert.equal(cell('INDEX({1,2;3,4},2,1)'), 3);
  assert.equal(cell('INDEX({1,2,3},2)'), 2);
  assert.deepEqual(array('INDEX({1,2;3,4},0,2)'), [[2], [4]]);
  assert.deepEqual(array('INDEX({1,2;3,4},1,0)'), [[1, 2]]);
  assert.equal(cell('INDEX(5,1,1)'), 5);
  assert.deepEqual(cell('INDEX(A1:B3,4,1)'), err('#REF!'));
  assert.deepEqual(cell('INDEX(A1:B3,1,3)'), err('#REF!'));
  assert.deepEqual(cell('INDEX({1,2},1,5)'), err('#REF!'));
  assert.deepEqual(cell('INDEX(A1:B3,-1,1)'), err('#VALUE!'));
  assert.deepEqual(cell('INDEX(A1:B3,1,1,2)'), err('#REF!'));
  assert.deepEqual(cell('INDEX(A1:B3,"x")'), err('#VALUE!'));
  assert.deepEqual(cell('INDEX(#N/A,1)'), err('#N/A'));
  assert.deepEqual(cell('INDEX(A1:B3)'), err('#VALUE!'));
});

test('information about cells', () => {
  assert.equal(cell('ISREF(A1)'), true);
  assert.equal(cell('ISREF(OFFSET(A1,1,1))'), true);
  assert.equal(cell('ISREF(1)'), false);
  assert.equal(cell('ISREF("A1")'), false);
  assert.equal(cell('ISREF(A1+0)'), false);
  assert.equal(cell('ISBLANK(Z99)'), true);
  assert.equal(cell('ISBLANK(A1)'), false);
  assert.equal(cell('ISBLANK(C4)'), false);
  assert.equal(cell('ISBLANK("")'), false);
  assert.deepEqual(array('ISBLANK(C3:C5)', 'Z100', ARRAY), [[false], [false], [true]]);
  assert.equal(cell('ISBLANK(C3:C5)', 'F5'), true);
  assert.equal(cell('ISBLANK(C3:C5)', 'F4'), false);
  assert.equal(cell('N(5)'), 5);
  assert.equal(cell('N(TRUE)'), 1);
  assert.equal(cell('N("7")'), 0);
  assert.equal(cell('N(A2:A3)'), 2);
  assert.deepEqual(cell('N(C3)'), err('#N/A'));
  assert.equal(cell('T("x")'), 'x');
  assert.equal(cell('T(5)'), '');
  assert.equal(cell('T(A4)'), 'text');
  assert.deepEqual(cell('T(#REF!)'), err('#REF!'));

  const formulas: Record<string, string> = { '0:0:0': 'B1*2', '1:1:1': 'SUM(A1)' };
  const options: EvaluateOptions = { cellFormula: (sheet, row, col) => formulas[`${sheet}:${row}:${col}`] };
  assert.equal(cell('ISFORMULA(A1)', 'Z100', options), true);
  assert.equal(cell('ISFORMULA(A2)', 'Z100', options), false);
  assert.equal(cell('ISFORMULA(A1)'), false);
  assert.deepEqual(cell('ISFORMULA(5)', 'Z100', options), err('#VALUE!'));
  assert.equal(cell('FORMULATEXT(A1)', 'Z100', options), '=B1*2');
  assert.equal(cell("FORMULATEXT('Other Sheet'!B2)", 'Z100', options), '=SUM(A1)');
  assert.deepEqual(cell('FORMULATEXT(A2)', 'Z100', options), err('#N/A'));

  assert.equal(cell('SHEET()'), 1);
  assert.equal(cell("SHEET('Other Sheet'!A1)"), 2);
  assert.equal(cell('SHEET("other sheet")'), 2);
  assert.deepEqual(cell('SHEET("nope")'), err('#N/A'));
  assert.equal(cell('SHEETS()'), 2);
  assert.equal(cell('SHEETS(A1:B2)'), 1);

  assert.equal(cell('CELL("row",B5)'), 5);
  assert.equal(cell('CELL("col",B5)'), 2);
  assert.equal(cell('CELL("ROW")', 'C7'), 7);
  assert.equal(cell('CELL("address",B5)'), '$B$5');
  assert.equal(cell('CELL("address",\'Other Sheet\'!B5)'), "'Other Sheet'!$B$5");
  assert.equal(cell('CELL("contents",B2)'), 20);
  assert.equal(cell('CELL("contents",Z99)'), 0);
  assert.deepEqual(cell('CELL("width",A1)'), err('#VALUE!'));
  assert.deepEqual(cell('CELL("row",5)'), err('#VALUE!'));

  assert.equal(cell('LET(x,1,x+1)'), 2);
  // A function is no value a cell can hold.
  assert.deepEqual(cell('LAMBDA(x,x+1)'), err('#CALC!'));
});

test('the built-in forms are never looked up in the library', () => {
  const trap = new Map<string, FunctionSpec>();
  const names = ['IF', 'IFERROR', 'IFNA', 'IFS', 'SWITCH', 'CHOOSE', 'AND', 'OR', 'XOR', 'NOT', 'ROW', 'COLUMN', 'ROWS', 'COLUMNS', 'AREAS', 'OFFSET', 'INDIRECT', 'INDEX', 'ISREF', 'ISFORMULA', 'FORMULATEXT', 'SHEET', 'SHEETS', 'ISBLANK', 'N', 'T', 'CELL', 'LET', 'LAMBDA'];
  for (const name of names) {
    assert.ok(isBuiltinForm(name), name);
    trap.set(name, { min: 0, call: () => 'library' });
  }
  assert.equal(isBuiltinForm('SUM'), false);
  const context = makeContext(SHEETS);
  for (const name of names) assert.notEqual(evaluate(parseFormula(`${name}(1,2)`), context, trap), 'library', name);
});

test('shared formulas: relative references move, absolute ones stay', () => {
  const node = parseFormula('A1+$B$1+A$1+$A1');
  const context = makeContext(SHEETS);
  assert.equal(toCellValue(evaluate(node, context, FUNCTIONS), context), 13);
  // One row down: A2 + $B$1 + A$1 + $A2.
  assert.equal(toCellValue(evaluate(node, context, FUNCTIONS, { rowShift: 1 }), context), 15);
  // One column right: B1 + $B$1 + B$1 + $A1.
  assert.equal(toCellValue(evaluate(node, context, FUNCTIONS, { colShift: 1 }), context), 31);
  assert.deepEqual(toCellValue(evaluate(node, context, FUNCTIONS, { rowShift: -1 }), context), err('#REF!'));
  assert.equal(toCellValue(evaluate(parseFormula('SUM(A:A)'), context, FUNCTIONS, { rowShift: 5, colShift: 1 }), context), 60);
  // Text given to INDIRECT and the formula of a name are not part of the shared formula.
  assert.equal(toCellValue(evaluate(parseFormula('INDIRECT("A1")+A1'), context, FUNCTIONS, { rowShift: 1 }), context), 3);
  const options: EvaluateOptions = { rowShift: 2, resolveName: () => parseFormula('A1') };
  assert.equal(toCellValue(evaluate(parseFormula('First+A1'), context, FUNCTIONS, options), context), 4);

  const ref = parseFormula('$A1:B$2');
  assert.ok(ref.type === 'ref');
  if (ref.type === 'ref') {
    assert.deepEqual(refRange(ref, 0, 0), { top: 0, left: 0, bottom: 1, right: 1 });
    assert.deepEqual(refRange(ref, 3, 2), { top: 1, left: 0, bottom: 3, right: 3 });
    assert.equal(refRange(ref, 0, 16384), undefined);
  }
});

test('long and deep formulas do not overflow the stack', () => {
  const terms: string[] = [];
  for (let i = 0; i < 2700; i++) terms.push('1');
  assert.equal(terms.join('+').length < 8192, true);
  assert.equal(cell(terms.join('+')), 2700);
  assert.equal(cell(terms.join('&')).toString().length, 2700);
  assert.equal(cell(`SUM(${terms.join(',')})`), 2700);
  assert.equal(cell('-'.repeat(200) + '5'), 5);
  assert.equal(cell('-'.repeat(201) + '5'), -5);
  assert.equal(cell('1' + '%'.repeat(3)), 0.000001);
  assert.equal(cell('('.repeat(250) + '1+1' + ')'.repeat(250)), 2);
  assert.equal(cell('SUM('.repeat(250) + '1' + ')'.repeat(250)), 1);
  assert.equal(cell('IF(1,'.repeat(250) + '7' + ')'.repeat(250)), 7);

  // A tree built by hand, far deeper than the parser would allow.
  const build = (depth: number, side: 'left' | 'right'): Node => {
    let node: Node = { type: 'number', value: 1 };
    for (let i = 0; i < depth; i++) {
      const one: Node = { type: 'number', value: 1 };
      node = side === 'left' ? { type: 'binary', op: '+', left: node, right: one } : { type: 'binary', op: '+', left: one, right: node };
    }
    return node;
  };
  const context = makeContext(SHEETS);
  assert.equal(evaluate(build(200000, 'left'), context, FUNCTIONS), 200001);
  // Deep on the right there is no way but recursion: it stops with an error instead of crashing.
  assert.deepEqual(evaluate(build(200000, 'right'), context, FUNCTIONS), err('#VALUE!'));
  let unary: Node = { type: 'number', value: 2 };
  for (let i = 0; i < 100001; i++) unary = { type: 'unary', op: '-', operand: unary };
  assert.equal(evaluate(unary, context, FUNCTIONS), -2);
  assert.equal(evaluationDepth(), 0);
  assert.equal(cell('DEPTH()'), 1);
  assert.equal(cell('SUM(SUM(DEPTH()))'), 3);
});

test('what a cell stores', () => {
  const context = makeContext(SHEETS, 'F2');
  assert.equal(toCellValue(5, context), 5);
  assert.equal(toCellValue(null, context), null);
  assert.equal(toCellValue({ rows: 2, cols: 1, data: ['a', 'b'] }, context), 'a');
  assert.deepEqual(toCellValue({ rows: 0, cols: 0, data: [] }, context), err('#VALUE!'));
  assert.equal(toCellValue({ ref: true, sheet: 0, range: { top: 1, left: 1, bottom: 1, right: 1 } }, context), 20);
  assert.equal(toCellValue({ ref: true, sheet: 0, range: { top: 0, left: 1, bottom: 2, right: 1 } }, context), 20);
  assert.equal(toNumber(toCellValue({ ref: true, sheet: 1, range: { top: 0, left: 0, bottom: 5, right: 0 } as CellRange }, context)), 0);
});
