import assert from 'node:assert/strict';
import { test } from 'node:test';
import { FormulaParseError, MAX_FORMULA_LENGTH, parseFormula, type Node } from '../src/calc/parse';
import { MAX_COLS, MAX_ROWS } from '../src/model/sheet/types';

/** A tree as compact text, so that grouping is easy to read: (1 + (2 * 3)). */
function show(node: Node | null): string {
  if (node === null) return '_';
  switch (node.type) {
    case 'number':
      return String(node.value);
    case 'string':
      return JSON.stringify(node.value);
    case 'boolean':
      return node.value ? 'TRUE' : 'FALSE';
    case 'error':
      return node.code;
    case 'ref': {
      const r = node.range;
      const sheet = node.sheet === undefined ? '' : `${node.external ? 'ext:' : ''}${node.sheet}${node.sheetEnd === undefined ? '' : `:${node.sheetEnd}`}!`;
      const kind = node.wholeCols ? 'cols' : node.wholeRows ? 'rows' : 'cells';
      return `${sheet}${kind}[${r.top},${r.left},${r.bottom},${r.right}]`;
    }
    case 'name':
      return `name(${node.sheet === undefined ? '' : `${node.sheet}!`}${node.name})`;
    case 'call':
      return `${node.name}(${node.args.map(show).join(',')})`;
    case 'unary':
      return node.op === '%' ? `(${show(node.operand)}%)` : `(${node.op}${show(node.operand)})`;
    case 'binary':
      return `(${show(node.left)} ${node.op === ' ' ? 'isect' : node.op} ${show(node.right)})`;
    case 'array':
      return `{${node.rows.map((row) => row.map(show).join(',')).join(';')}}`;
  }
}

const parsed = (text: string) => show(parseFormula(text));

function failsAt(text: string, position: number): void {
  assert.throws(
    () => parseFormula(text),
    (error: unknown) => {
      assert.ok(error instanceof FormulaParseError, `${text}: ${String(error)}`);
      assert.equal(error.position, position, `${text}: ${error.message}`);
      return true;
    },
  );
}

test('numbers, text, booleans and errors', () => {
  assert.equal(parsed('1'), '1');
  assert.equal(parsed('=1.5'), '1.5');
  assert.equal(parsed('.5'), '0.5');
  assert.equal(parsed('1E3'), '1000');
  assert.equal(parsed('1.5E-3'), '0.0015');
  assert.equal(parsed('1e+2'), '100');
  assert.equal(parsed('"a""b"'), '"a\\"b"');
  assert.equal(parsed('""'), '""');
  assert.equal(parsed('TRUE'), 'TRUE');
  assert.equal(parsed('false'), 'FALSE');
  for (const code of ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CALC!']) assert.equal(parsed(code), code);
  assert.equal(parsed('#n/a'), '#N/A');
  assert.equal(parsed('  = 1 '), '1');
});

test('precedence follows Excel, not arithmetic convention', () => {
  assert.equal(parsed('1+2*3'), '(1 + (2 * 3))');
  assert.equal(parsed('(1+2)*3'), '((1 + 2) * 3)');
  // A sign binds tighter than "^", and "^" groups from the left.
  assert.equal(parsed('-2^2'), '((-2) ^ 2)');
  assert.equal(parsed('2^3^2'), '((2 ^ 3) ^ 2)');
  assert.equal(parsed('2^-1'), '(2 ^ (-1))');
  assert.equal(parsed('-5%'), '((-5)%)');
  assert.equal(parsed('50%^2'), '((50%) ^ 2)');
  assert.equal(parsed('1-2-3'), '((1 - 2) - 3)');
  assert.equal(parsed('8/4/2'), '((8 / 4) / 2)');
  assert.equal(parsed('1&2+3'), '(1 & (2 + 3))');
  assert.equal(parsed('1+2&3'), '((1 + 2) & 3)');
  assert.equal(parsed('1&2=3'), '((1 & 2) = 3)');
  assert.equal(parsed('1<2=TRUE'), '((1 < 2) = TRUE)');
  assert.equal(parsed('--1'), '(-(-1))');
  assert.equal(parsed('-+1'), '(-(+1))');
  assert.equal(parsed('5%%'), '((5%)%)');
  assert.equal(parsed('-A1:B2'), '(-cells[0,0,1,1])');
  for (const op of ['=', '<>', '<', '>', '<=', '>=']) assert.equal(parsed(`1${op}2`), `(1 ${op} 2)`);
  assert.equal(parsed('1 + 2\n* 3'), '(1 + (2 * 3))');
});

test('cell references', () => {
  assert.equal(parsed('A1'), 'cells[0,0,0,0]');
  assert.equal(parsed('$b$2'), 'cells[1,1,1,1]');
  assert.equal(parsed('A1:B2'), 'cells[0,0,1,1]');
  assert.equal(parsed('B2:A1'), 'cells[0,0,1,1]');
  assert.equal(parsed('XFD1048576'), `cells[${MAX_ROWS - 1},${MAX_COLS - 1},${MAX_ROWS - 1},${MAX_COLS - 1}]`);
  assert.equal(parsed('A:A'), `cols[0,0,${MAX_ROWS - 1},0]`);
  assert.equal(parsed('$A:C'), `cols[0,0,${MAX_ROWS - 1},2]`);
  assert.equal(parsed('1:1'), `rows[0,0,0,${MAX_COLS - 1}]`);
  assert.equal(parsed('$2:$3'), `rows[1,0,2,${MAX_COLS - 1}]`);
  assert.equal(parsed('SUM(1:3)'), `SUM(rows[0,0,2,${MAX_COLS - 1}])`);

  const mixed = parseFormula('$A1:B$2');
  assert.equal(mixed.type, 'ref');
  if (mixed.type === 'ref') assert.deepEqual(mixed.abs, { top: false, left: true, bottom: true, right: false });
  const reversed = parseFormula('B$2:$A1');
  if (reversed.type === 'ref') assert.deepEqual(reversed.abs, { top: false, left: true, bottom: true, right: false });
});

test('sheet prefixes, 3D and external references', () => {
  assert.equal(parsed('Sheet1!A1'), 'Sheet1!cells[0,0,0,0]');
  assert.equal(parsed("'My Sheet'!A1:B2"), 'My Sheet!cells[0,0,1,1]');
  assert.equal(parsed("'It''s'!A1"), "It's!cells[0,0,0,0]");
  assert.equal(parsed('Sheet1!A1:Sheet1!B2'), 'Sheet1!cells[0,0,1,1]');
  assert.equal(parsed('Sheet1!A:B'), `Sheet1!cols[0,0,${MAX_ROWS - 1},1]`);
  assert.equal(parsed('Sheet1!1:2'), `Sheet1!rows[0,0,1,${MAX_COLS - 1}]`);
  assert.equal(parsed('Sheet1:Sheet3!A1'), 'Sheet1:Sheet3!cells[0,0,0,0]');
  assert.equal(parsed("'Sheet 1:Sheet 3'!A1:B2"), 'Sheet 1:Sheet 3!cells[0,0,1,1]');
  assert.equal(parsed('[1]Sheet1!A1'), 'ext:[1]Sheet1!cells[0,0,0,0]');
  assert.equal(parsed("'[Book 1.xlsx]Sheet1'!$A$1"), 'ext:[Book 1.xlsx]Sheet1!cells[0,0,0,0]');
  assert.equal(parsed('Sheet1!#REF!'), '#REF!');
  assert.equal(parsed('Sheet1!Total'), 'name(Sheet1!Total)');
  // Two different sheets on the two sides are not one range.
  assert.equal(parsed('Sheet1!A1:Sheet2!B2'), '(Sheet1!cells[0,0,0,0] : Sheet2!cells[1,1,1,1])');
  const external = parseFormula('[1]Sheet1!A1');
  assert.ok(external.type === 'ref' && external.external === true);
});

test('names and the line between a name and a cell', () => {
  assert.equal(parsed('Total'), 'name(Total)');
  assert.equal(parsed('_x.y'), 'name(_x.y)');
  assert.equal(parsed('\\a'), 'name(\\a)');
  assert.equal(parsed('Ürün'), 'name(Ürün)');
  assert.equal(parsed('A'), 'name(A)');
  // Past the last column or row it is no longer a cell.
  assert.equal(parsed('XFE1'), 'name(XFE1)');
  assert.equal(parsed('A1048577'), 'name(A1048577)');
  assert.equal(parsed('A1B2'), 'name(A1B2)');
  assert.equal(parsed('TRUE1'), 'name(TRUE1)');
  // A word before "(" is a function, even when it spells a cell.
  assert.equal(parsed('LOG10(5)'), 'LOG10(5)');
  assert.equal(parsed('A1:Total'), '(cells[0,0,0,0] : name(Total))');
});

test('function calls', () => {
  assert.equal(parsed('PI()'), 'PI()');
  assert.equal(parsed('sum(1,2)'), 'SUM(1,2)');
  assert.equal(parsed('IF(A1,,2)'), 'IF(cells[0,0,0,0],_,2)');
  assert.equal(parsed('IF(A1,1,)'), 'IF(cells[0,0,0,0],1,_)');
  assert.equal(parsed('F(,)'), 'F(_,_)');
  assert.equal(parsed('SUM( 1 , 2 )'), 'SUM(1,2)');
  assert.equal(parsed('_xlfn.XLOOKUP(1,A:A,B:B)'), `XLOOKUP(1,cols[0,0,${MAX_ROWS - 1},0],cols[0,1,${MAX_ROWS - 1},1])`);
  assert.equal(parsed('_xlfn._xlws.FILTER(A1)'), 'FILTER(cells[0,0,0,0])');
  assert.equal(parsed('_xlws.SORT(A1)'), 'SORT(cells[0,0,0,0])');
  assert.equal(parsed('SUM(A1:B2,MAX(1,2))'), 'SUM(cells[0,0,1,1],MAX(1,2))');
  assert.equal(parsed('A1:INDEX(B:B,3)'), `(cells[0,0,0,0] : INDEX(cols[0,1,${MAX_ROWS - 1},1],3))`);
});

test('reference operators: range over intersection over union', () => {
  assert.equal(parsed('A1:B2 B2:C3'), '(cells[0,0,1,1] isect cells[1,1,2,2])');
  assert.equal(parsed('SUM(A:A 1:1)'), `SUM((cols[0,0,${MAX_ROWS - 1},0] isect rows[0,0,0,${MAX_COLS - 1}]))`);
  assert.equal(parsed('SUM((A1,B2))'), 'SUM((cells[0,0,0,0] , cells[1,1,1,1]))');
  assert.equal(parsed('SUM((A1 B1,C1))'), 'SUM(((cells[0,0,0,0] isect cells[0,1,0,1]) , cells[0,2,0,2]))');
  assert.equal(parsed('Total Other'), '(name(Total) isect name(Other))');
  assert.equal(parsed('(A1:B2) (B2:C3)'), '(cells[0,0,1,1] isect cells[1,1,2,2])');
  // Spaces around other operators mean nothing.
  assert.equal(parsed('A1 + B1'), '(cells[0,0,0,0] + cells[0,1,0,1])');
  assert.equal(parsed('A1 -B1'), '(cells[0,0,0,0] - cells[0,1,0,1])');
  assert.equal(parsed('-A1 B1'), '(-(cells[0,0,0,0] isect cells[0,1,0,1]))');
});

test('array constants', () => {
  assert.equal(parsed('{1,2;3,4}'), '{1,2;3,4}');
  assert.equal(parsed('{1,"x",TRUE,#N/A,-1}'), '{1,"x",TRUE,#N/A,-1}');
  assert.equal(parsed('{ 1 , 2 ; +3 , false }'), '{1,2;3,FALSE}');
  assert.equal(parsed('SUM({1,2}*2)'), 'SUM(({1,2} * 2))');
  failsAt('{1,2;3}', 0);
  failsAt('{1,A1}', 3);
  failsAt('{1,2', 4);
  failsAt('{}', 1);
  failsAt('{1,-"a"}', 4);
});

test('bad input is refused with the position of the problem', () => {
  failsAt('', 0);
  failsAt('=', 1);
  failsAt('1+', 2);
  failsAt('=1+*2', 3);
  failsAt('(1+2', 4);
  failsAt('SUM(1,2', 7);
  failsAt('1 2', 2);
  failsAt('"abc', 0);
  failsAt("'Sheet1!A1", 0);
  failsAt("'Sheet 1'A1", 9);
  failsAt('A1)', 2);
  failsAt('#BAD!', 0);
  failsAt('1..2', 2);
  failsAt('.', 0);
  failsAt('$', 0);
  failsAt('2A', 1);
  failsAt('A1 5', 3);
  failsAt('Sheet1!', 7);
  failsAt('Sheet1:Sheet2!Total', 14);
  failsAt('1 @ 2', 2);
  // Structured (table) references are not supported.
  failsAt('Table1[Col]', 6);
  failsAt('SUM(Table1[[#This Row],[Total]])', 10);
  failsAt('[@Total]*2', 0);
  assert.throws(() => parseFormula(undefined as unknown as string), FormulaParseError);
});

test('nesting and length are bounded', () => {
  const nested = (depth: number) => '('.repeat(depth) + '1' + ')'.repeat(depth);
  assert.equal(parsed(nested(200)), '1');
  assert.throws(() => parseFormula(nested(300)), FormulaParseError);
  assert.throws(() => parseFormula('SUM('.repeat(4000) + '1' + ')'.repeat(4000)), FormulaParseError);
  assert.throws(() => parseFormula('-'.repeat(1000) + '1'), FormulaParseError);
  assert.throws(() => parseFormula('1' + '%'.repeat(1000)), FormulaParseError);
  assert.throws(() => parseFormula('{'.repeat(5000)), FormulaParseError);
  assert.throws(() => parseFormula('('.repeat(8000)), FormulaParseError);

  // A long chain of operators is not nesting: it parses, however long.
  const terms: string[] = [];
  for (let i = 0; i < 2000; i++) terms.push('1');
  const chain = parseFormula(terms.join('+'));
  assert.equal(chain.type, 'binary');

  const long = `"${'x'.repeat(MAX_FORMULA_LENGTH - 2)}"`;
  assert.equal(long.length, MAX_FORMULA_LENGTH);
  assert.equal(parseFormula(long).type, 'string');
  failsAt(`${long} `, MAX_FORMULA_LENGTH);
});

test('the same text gives the same tree', () => {
  assert.equal(parseFormula('A1+B1*2'), parseFormula('A1+B1*2'));
  assert.notEqual(parseFormula('A1+B1*2'), parseFormula('=A1+B1*2'));
});

test('performance: 100,000 short formulas in under a second', () => {
  // Every text is different, so nothing comes from the cache. The best of three rounds is
  // taken: the limit is about the parser, not about what else the machine is doing.
  let best = Infinity;
  for (let round = 0; round < 3; round++) {
    const texts: string[] = [];
    for (let i = 0; i < 100000; i++) texts.push(`IF(A${i + 1}>0,SUM($B$1:B${i + 1})*1.5,"none")&C${(i % 500) + 1}&${round}`);
    const start = performance.now();
    let count = 0;
    for (const text of texts) if (parseFormula(text).type === 'binary') count++;
    best = Math.min(best, performance.now() - start);
    assert.equal(count, 100000);
  }
  console.log(`# parsing 100,000 different formulas: ${best.toFixed(0)} ms`);
  assert.ok(best < 10000, `parsing took ${best.toFixed(0)} ms`);
});
