// Dynamic arrays (spilling, "#", "@"), LET, LAMBDA and the functions that take a LAMBDA.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine, type CellAddress } from '../src/calc/engine';
import { needsDynamic } from '../src/calc/dynamic';
import { parseFormula } from '../src/calc/parse';
import { isError, type Scalar } from '../src/calc/values';
import { createSheet, emptyStyleSheet, getCell, parseCellRef, type Cell, type CellFormula, type DefinedName, type Row, type Sheet, type Workbook } from '../src/model/sheet/types';

type Plain = number | string | boolean | null;
/** "=..." is a formula typed into the cell (dynamic where needsDynamic() says so); "~=..." an ordinary one of a file. */
type Spec = Plain | { f: string; v?: Plain; dynamic?: boolean; ref?: string; error?: boolean };

const emptyRow = (): Row => ({ customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] });

function makeCell(spec: Spec): Cell {
  if (typeof spec === 'string' && spec.charAt(0) === '=') spec = { f: spec.slice(1), dynamic: needsDynamic(spec.slice(1)) };
  else if (typeof spec === 'string' && spec.startsWith('~=')) spec = { f: spec.slice(2) };
  if (spec !== null && typeof spec === 'object') {
    const value = spec.v === undefined ? null : spec.v;
    const formula: CellFormula = { text: spec.f, type: 'normal' };
    if (spec.dynamic) formula.dynamic = true;
    if (spec.ref !== undefined) formula.ref = spec.ref;
    const kind: Cell['kind'] = spec.error ? 'error' : value === null ? 'blank' : typeof value === 'number' ? 'number' : typeof value === 'boolean' ? 'boolean' : 'string';
    return { kind, value, style: 0, formula };
  }
  const kind: Cell['kind'] = spec === null ? 'blank' : typeof spec === 'number' ? 'number' : typeof spec === 'boolean' ? 'boolean' : 'string';
  return { kind, value: spec, style: 0 };
}

function place(sheet: Sheet, address: string, cell: Cell | undefined): CellAddress {
  const at = parseCellRef(address) as { row: number; col: number };
  const old = sheet.rows.get(at.row);
  const row: Row = old ? { ...old, cells: old.cells.slice() } : emptyRow();
  if (cell) row.cells[at.col] = cell;
  else delete row.cells[at.col];
  sheet.rows.set(at.row, row);
  if (at.row >= sheet.rowCount) sheet.rowCount = at.row + 1;
  if (at.col >= sheet.colCount) sheet.colCount = at.col + 1;
  return { sheet: 0, row: at.row, col: at.col };
}

function book(cells: Record<string, Spec>, definedNames: DefinedName[] = []): Workbook {
  const workbook: Workbook = { sheets: [], styles: emptyStyleSheet(), sharedStrings: [], definedNames, date1904: false, activeSheet: 0 };
  const sheet = createSheet({ name: 'S', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  for (const address of Object.keys(cells)) place(sheet, address, makeCell(cells[address]));
  workbook.sheets.push(sheet);
  return workbook;
}

function shown(workbook: Workbook, address: string): Scalar | undefined {
  const at = parseCellRef(address) as { row: number; col: number };
  const cell = getCell(workbook.sheets[0], at.row, at.col);
  if (!cell) return undefined;
  if (cell.kind === 'error') return `${cell.value}`;
  return cell.value;
}

function column(workbook: Workbook, from: string, count: number): (Scalar | undefined)[] {
  const at = parseCellRef(from) as { row: number; col: number };
  const out: (Scalar | undefined)[] = [];
  for (let k = 0; k < count; k++) {
    const cell = getCell(workbook.sheets[0], at.row + k, at.col);
    out.push(cell ? (cell.kind === 'error' ? String(cell.value) : cell.value) : undefined);
  }
  return out;
}

/** Calculates every formula, as for a workbook just typed in. */
function calculate(cells: Record<string, Spec>, names: DefinedName[] = []): { workbook: Workbook; engine: CalcEngine } {
  const workbook = book(cells, names);
  const engine = new CalcEngine(workbook, { budget: { milliseconds: 0 } });
  engine.recalculate(engine.uncalculated());
  return { workbook, engine };
}

function value(formula: string, cells: Record<string, Spec> = {}, names: DefinedName[] = []): Scalar | undefined {
  const { workbook } = calculate({ ...cells, Z1: `=${formula}` }, names);
  return shown(workbook, 'Z1');
}

// ---------------------------------------------------------------------------
// Syntax

test('"#", "@" and a call of a call are read as the calls files store', () => {
  assert.deepEqual(parseFormula('A1#'), { type: 'call', name: 'ANCHORARRAY', args: [parseFormula('A1')] });
  assert.deepEqual(parseFormula('_xlfn.ANCHORARRAY(A1)'), parseFormula('A1#'));
  assert.deepEqual(parseFormula('@A1:A3'), { type: 'call', name: 'SINGLE', args: [parseFormula('A1:A3')] });
  assert.deepEqual(parseFormula('_xlfn.SINGLE(A1:A3)'), parseFormula('@A1:A3'));
  const applied = parseFormula('LAMBDA(x,x+1)(2)');
  assert.equal(applied.type, 'call');
  assert.equal((applied as { name: string }).name, '()');
  assert.deepEqual(parseFormula('_xlfn.LET(_xlpm.x,1,_xlpm.x+1)'), parseFormula('LET(x,1,x+1)'));
  assert.deepEqual(parseFormula('SUM(Sheet2!B3#)'), { type: 'call', name: 'SUM', args: [{ type: 'call', name: 'ANCHORARRAY', args: [parseFormula('Sheet2!B3')] }] });
  // "A1:B2#" is a range to the spill of B2; "#" after a whole column or a number is no spill reference.
  assert.equal((parseFormula('A1:B2#') as { op?: string }).op, ':');
  assert.throws(() => parseFormula('A:A#'));
  assert.throws(() => parseFormula('1#'));
});

// ---------------------------------------------------------------------------
// LET and LAMBDA

test('LET names values for what follows', () => {
  assert.equal(value('LET(x,2,y,x*3,x+y)'), 8);
  assert.equal(value('LET(x,A1:A3,SUM(x))', { A1: 1, A2: 2, A3: 4 }), 7);
  // A name of LET hides a defined name of the same spelling.
  assert.equal(value('LET(Rate,5,Rate*2)', {}, [{ name: 'Rate', value: '100' }]), 10);
  assert.equal(value('LET(x,1)'), '#VALUE!');
  assert.equal(value('LET(A1,1,2)'), '#VALUE!');
  // A name is seen only after it is given.
  assert.equal(value('LET(x,y,y,1,x)'), '#NAME?');
});

test('LAMBDA is called where it is written, through LET and through a defined name', () => {
  assert.equal(value('LAMBDA(x,x+1)(2)'), 3);
  assert.equal(value('LAMBDA(x,y,x*y)(3,4)'), 12);
  assert.equal(value('LET(f,LAMBDA(x,x*x),f(5)+f(2))'), 29);
  const names = [{ name: 'Topla2', value: 'LAMBDA(a,b,a+b)' }];
  assert.equal(value('Topla2(3,4)', {}, names), 7);
  assert.equal(value('Topla2(A1,10)', { A1: 5 }, names), 15);
  // A function is no value of a cell.
  assert.equal(value('LAMBDA(x,x+1)'), '#CALC!');
  assert.equal(value('Topla2', {}, names), '#CALC!');
  assert.equal(value('LET(f,LAMBDA(x,x),f)'), '#CALC!');
  assert.equal(value('SUM(LAMBDA(x,x))'), '#CALC!');
  // More arguments than parameters; the same parameter twice.
  assert.equal(value('LAMBDA(x,x)(1,2)'), '#VALUE!');
  assert.equal(value('LAMBDA(x,x,x)(1,2)'), '#VALUE!');
  // A LAMBDA sees the names where it was made, not where it is called.
  assert.equal(value('LET(k,10,f,LAMBDA(x,x+k),LET(k,1000,f(1)))'), 11);
  // Currying: a LAMBDA that gives a LAMBDA.
  assert.equal(value('LAMBDA(a,LAMBDA(b,a-b))(10)(3)'), 7);
});

test('a parameter left out is omitted', () => {
  assert.equal(value('LAMBDA(x,y,IF(ISOMITTED(y),x,x+y))(1)'), 1);
  assert.equal(value('LAMBDA(x,y,IF(ISOMITTED(y),x,x+y))(1,2)'), 3);
  assert.equal(value('LAMBDA(x,y,IF(ISOMITTED(y),"none",y))(1,)'), 'none');
});

test('a defined name may call itself, down to a limit', () => {
  const names = [{ name: 'Fact', value: 'LAMBDA(n,IF(n<=1,1,n*Fact(n-1)))' }];
  assert.equal(value('Fact(5)', {}, names), 120);
  assert.equal(value('Fact(10)', {}, names), 3628800);
  const deep = value('Fact(100000)', {}, names);
  assert.ok(isError({ error: String(deep) }) && (deep === '#NUM!' || deep === '#VALUE!'), `deep recursion gives an error, not ${deep}`);
  // Two names that call each other.
  const pair = [
    { name: 'IsEven', value: 'LAMBDA(n,IF(n=0,TRUE,IsOdd(n-1)))' },
    { name: 'IsOdd', value: 'LAMBDA(n,IF(n=0,FALSE,IsEven(n-1)))' },
  ];
  assert.equal(value('IsEven(10)', {}, pair), true);
  assert.equal(value('IsOdd(7)', {}, pair), true);
});

test('a long chain of names that stand for one another ends in an error, not in a stack overflow', () => {
  const names = Array.from({ length: 20000 }, (_, k) => ({ name: `Step_${k}`, value: k === 19999 ? "LAMBDA(x,x+1)" : `Step_${k + 1}` }));
  assert.match(String(value("Step_0(1)", {}, names)), /^#(VALUE|NUM)!$/);
  assert.match(String(value("LET(f,Step_0,f(1))", {}, names)), /^#(VALUE|NUM)!$/);
  // A short one works.
  assert.equal(value("Step_19990(1)", {}, names), 2);
});

test('a LAMBDA that does not end runs out of budget instead of hanging', () => {
  const names = [{ name: 'Fib', value: 'LAMBDA(n,IF(n<2,n,Fib(n-1)+Fib(n-2)))' }];
  const workbook = book({ A1: '=Fib(60)' }, names);
  const engine = new CalcEngine(workbook, { budget: { milliseconds: 50, hardMilliseconds: 200 } });
  const started = Date.now();
  for (let k = 0; k < 20 && (k === 0 || engine.incomplete); k++) engine.recalculate(k === 0 ? engine.uncalculated() : []);
  assert.ok(Date.now() - started < 10000);
  assert.equal(engine.overBudget.length, 1);
});

test('MAP, REDUCE, SCAN, BYROW, BYCOL and MAKEARRAY', () => {
  const cells = { A1: 1, A2: 2, A3: 3, B1: 10, B2: 20, B3: 30 };
  const { workbook } = calculate({
    ...cells,
    D1: '=MAP(A1:A3,LAMBDA(x,x*2))',
    E1: '=MAP(A1:A3,B1:B3,LAMBDA(x,y,x+y))',
    F1: '=REDUCE(0,A1:B3,LAMBDA(a,v,a+v))',
    G1: '=SCAN(0,A1:A3,LAMBDA(a,v,a+v))',
    H1: '=BYROW(A1:B3,LAMBDA(r,SUM(r)))',
    I1: '=BYCOL(A1:B3,LAMBDA(c,MAX(c)))',
    K1: '=MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))',
  });
  assert.deepEqual(column(workbook, 'D1', 3), [2, 4, 6]);
  assert.deepEqual(column(workbook, 'E1', 3), [11, 22, 33]);
  assert.equal(shown(workbook, 'F1'), 66);
  assert.deepEqual(column(workbook, 'G1', 3), [1, 3, 6]);
  assert.deepEqual(column(workbook, 'H1', 3), [11, 22, 33]);
  assert.deepEqual([shown(workbook, 'I1'), shown(workbook, 'J1')], [3, 30]);
  assert.deepEqual([shown(workbook, 'K1'), shown(workbook, 'L1'), shown(workbook, 'M1'), shown(workbook, 'K2'), shown(workbook, 'M2')], [11, 12, 13, 21, 23]);
  // An array where one value is wanted; a LAMBDA with the wrong number of parameters.
  assert.equal(value('MAP(A1:A3,LAMBDA(x,SEQUENCE(2)))', cells), '#CALC!');
  assert.equal(value('MAP(A1:A3,LAMBDA(x,y,x))', cells), '#VALUE!');
  assert.equal(value('MAKEARRAY(0,1,LAMBDA(r,c,1))'), '#VALUE!');
  assert.equal(value('MAKEARRAY(100000,100000,LAMBDA(r,c,1))'), '#NUM!');
  // REDUCE may build an array.
  assert.equal(value('ROWS(REDUCE(0,A1:A3,LAMBDA(a,v,SEQUENCE(v))))', cells), 3);
});

// ---------------------------------------------------------------------------
// Spilling

test('a dynamic array formula spills into the cells below and to the right', () => {
  const { workbook } = calculate({ A1: '=SEQUENCE(3,2)' });
  assert.deepEqual([shown(workbook, 'A1'), shown(workbook, 'B1'), shown(workbook, 'A2'), shown(workbook, 'B3')], [1, 2, 3, 6]);
  const anchor = getCell(workbook.sheets[0], 0, 0) as Cell;
  assert.equal(anchor.formula && anchor.formula.ref, 'A1:B3');
  // The cells it fills hold values, no formulas.
  assert.equal((getCell(workbook.sheets[0], 2, 1) as Cell).formula, undefined);
});

test('ordinary expressions spill too: a range times a number, a range', () => {
  const cells = { A1: 1, A2: 2, A3: 3 };
  const { workbook } = calculate({ ...cells, C1: '=A1:A3*2', D1: '=A1:A3' });
  assert.deepEqual(column(workbook, 'C1', 3), [2, 4, 6]);
  assert.deepEqual(column(workbook, 'D1', 3), [1, 2, 3]);
  // An ordinary formula of a file takes one cell of the range instead (implicit intersection).
  const legacy = calculate({ ...cells, C2: '~=A1:A3*2' }).workbook;
  assert.equal(shown(legacy, 'C2'), 4);
  assert.equal(shown(legacy, 'C3'), undefined);
});

test('an ordinary formula calculated for a dynamic one that reads it keeps its implicit intersection', () => {
  // D1 comes first in the row order and needs B2 while it is being calculated.
  const { workbook } = calculate({ A1: 1, A2: 2, A3: 3, B2: '~=A1:A3*2', D1: '=B2*10', E1: '=SEQUENCE(2)', F1: '=E1+B2' });
  assert.equal(shown(workbook, 'B2'), 4);
  assert.equal(shown(workbook, 'D1'), 40);
  assert.equal((getCell(workbook.sheets[0], 0, 3) as Cell).formula?.ref, undefined, 'D1 reads one cell: an ordinary formula would mean the same');
  assert.equal(shown(workbook, 'F1'), 5);
});

test('"@" asks for implicit intersection in a dynamic array formula', () => {
  const { workbook } = calculate({ A1: 1, A2: 2, A3: 3, C2: '=@A1:A3*2', D1: '=SUM(@A1:A3)' });
  assert.equal(shown(workbook, 'C2'), 4);
  assert.equal(shown(workbook, 'C3'), undefined);
  assert.equal(shown(workbook, 'D1'), 1);
});

test('FILTER, SORT, SORTBY, UNIQUE and XLOOKUP spill their results', () => {
  const cells = { A1: 'b', A2: 'a', A3: 'c', A4: 'a', B1: 3, B2: 1, B3: 2, B4: 5 };
  const { workbook } = calculate({
    ...cells,
    D1: '=FILTER(A1:A4,B1:B4>1)',
    E1: '=SORT(A1:A4)',
    F1: '=SORTBY(A1:A4,B1:B4,-1)',
    G1: '=UNIQUE(A1:A4)',
    H1: '=XLOOKUP("c",A1:A4,A1:B4)',
  });
  assert.deepEqual(column(workbook, 'D1', 4), ['b', 'c', 'a', undefined]);
  assert.deepEqual(column(workbook, 'E1', 4), ['a', 'a', 'b', 'c']);
  assert.deepEqual(column(workbook, 'F1', 4), ['a', 'b', 'c', 'a']);
  assert.deepEqual(column(workbook, 'G1', 4), ['b', 'a', 'c', undefined]);
  assert.deepEqual([shown(workbook, 'H1'), shown(workbook, 'I1')], ['c', 2]);
});

test('something in the way makes the formula #SPILL!, and it spills once it is gone', () => {
  const workbook = book({ A1: '=SEQUENCE(3)', A3: 'x' });
  const engine = new CalcEngine(workbook);
  engine.recalculate(engine.uncalculated());
  assert.equal(shown(workbook, 'A1'), '#SPILL!');
  assert.equal(shown(workbook, 'A2'), undefined);
  assert.equal(shown(workbook, 'A3'), 'x');
  // The formula keeps its own cell as where it goes.
  assert.equal((getCell(workbook.sheets[0], 0, 0) as Cell).formula?.ref, 'A1');
  engine.recalculate([place(workbook.sheets[0], 'A3', undefined)]);
  assert.deepEqual(column(workbook, 'A1', 3), [1, 2, 3]);
  // Typing into a cell of the spill: the formula is #SPILL! and gives the other cells up.
  engine.recalculate([place(workbook.sheets[0], 'A2', makeCell(7))]);
  assert.deepEqual(column(workbook, 'A1', 3), ['#SPILL!', 7, undefined]);
  // A formula in the way blocks it as well, even one that gives nothing.
  engine.recalculate([place(workbook.sheets[0], 'A2', makeCell('=""'))]);
  assert.equal(shown(workbook, 'A1'), '#SPILL!');
  engine.recalculate([place(workbook.sheets[0], 'A2', undefined)]);
  assert.deepEqual(column(workbook, 'A1', 3), [1, 2, 3]);
});

test('a merged cell and the edge of the sheet are in the way', () => {
  const workbook = book({ A1: '=SEQUENCE(3)' });
  workbook.sheets[0].merges = [{ top: 2, left: 0, bottom: 2, right: 1 }];
  const engine = new CalcEngine(workbook);
  engine.recalculate(engine.uncalculated());
  assert.equal(shown(workbook, 'A1'), '#SPILL!');
  assert.equal(value('SEQUENCE(2000000)'), '#SPILL!');
  // Too large to make at all is another error.
  assert.equal(value('SEQUENCE(1000000000)'), '#NUM!');
});

test('large spills are written in about the time an array formula of the same size takes, and the workbook has a budget for them', () => {
  // Two formulas share the budget of 2 million cells: 800,000 each fit.
  const workbook = book({ A1: '=SEQUENCE(800000)', B1: '=SEQUENCE(800000)' });
  const engine = new CalcEngine(workbook, { budget: { milliseconds: 250 } });
  const started = Date.now();
  engine.recalculate(engine.uncalculated());
  for (let k = 0; k < 200 && engine.incomplete; k++) engine.recalculate([]);
  const took = Date.now() - started;
  assert.equal(engine.incomplete, false);
  const sheet = workbook.sheets[0];
  assert.deepEqual([0, 1].map((col) => getCell(sheet, 799999, col)?.value), [800000, 800000]);
  // A guard against work that grows faster than the cells (before: tens of seconds), not a measure of speed.
  assert.ok(took < 30000, `${took} ms`);
  // A third one, after them on the sheet, does not fit into what is left of the 2 million: it is #SPILL!.
  engine.recalculate([place(sheet, 'C1', makeCell('=SEQUENCE(800000)'))]);
  for (let k = 0; k < 200 && engine.incomplete; k++) engine.recalculate([]);
  assert.deepEqual([0, 1, 2].map((col) => `${getCell(sheet, 0, col)?.value}`), ['1', '1', '#SPILL!']);
  assert.equal(getCell(sheet, 799999, 2), undefined);
  // It goes: the two spill as before.
  engine.recalculate([place(sheet, 'C1', undefined)]);
  for (let k = 0; k < 200 && engine.incomplete; k++) engine.recalculate([]);
  assert.deepEqual([0, 1].map((col) => getCell(sheet, 799999, col)?.value), [800000, 800000]);
});

test('the budget of spills goes by place, first come first served on the sheet, whatever the order of calculating (K-E)', () => {
  const outcomes = new Set<string>();
  for (const order of [['A1', 'B1', 'C1'], ['C1', 'B1', 'A1'], ['B1', 'C1', 'A1']]) {
    const workbook = book({ G1: 1, A1: '=SEQUENCE(G1)', B1: '=SEQUENCE(G1)', C1: '=SEQUENCE(3)' });
    const engine = new CalcEngine(workbook, { limits: { workbookSpillCells: 12 } });
    engine.recalculate(engine.uncalculated());
    // 5 + 5 + 3 cells wanted, 12 to share: A1 and B1 (first by place) get theirs, C1 is #SPILL!.
    engine.recalculate([place(workbook.sheets[0], 'G1', makeCell(5)), ...order.map((a) => place(workbook.sheets[0], a, makeCell(`=SEQUENCE(${a === 'C1' ? 3 : 'G1'})`)))]);
    outcomes.add(['A1', 'B1', 'C1'].map((a) => `${shown(workbook, a)}`).join(','));
    engine.recalculate([place(workbook.sheets[0], 'G1', makeCell(4))]);
    outcomes.add(['A1', 'B1', 'C1'].map((a) => `${shown(workbook, a)}`).join(','));
  }
  assert.deepEqual(Array.from(outcomes), ['1,1,#SPILL!', '1,1,1']);
});

test('formulas of one value do not count in the budget of spills, and typing them in any order changes nothing (K-E)', () => {
  for (const first of ['spill', 'singles']) {
    const workbook = book({ B1: 1 });
    const engine = new CalcEngine(workbook, { limits: { workbookSpillCells: 20 } });
    const typeSpill = () => engine.recalculate([place(workbook.sheets[0], 'A1', makeCell('=SEQUENCE(15)'))]);
    const typeSingles = () => {
      const cells = [];
      for (let r = 1; r <= 200; r++) cells.push(place(workbook.sheets[0], `F${r}`, makeCell('=INDEX(B1:B5,1)')));
      engine.recalculate(cells);
    };
    if (first === 'spill') {
      typeSpill();
      typeSingles();
    } else {
      typeSingles();
      typeSpill();
    }
    assert.deepEqual([shown(workbook, 'A1'), shown(workbook, 'A15'), shown(workbook, 'F1')], [1, 15, 1], first);
  }
});

test('a spill that grows or shrinks updates the cells and what reads them', () => {
  const workbook = book({ A1: 3, B1: '=SEQUENCE(A1)', C1: '=SUM(B1:B10)', D1: '=B5', E1: '=COUNT(B1#)' });
  const engine = new CalcEngine(workbook);
  engine.recalculate(engine.uncalculated());
  assert.deepEqual(column(workbook, 'B1', 5), [1, 2, 3, undefined, undefined]);
  assert.deepEqual([shown(workbook, 'C1'), shown(workbook, 'D1'), shown(workbook, 'E1')], [6, 0, 3]);
  engine.recalculate([place(workbook.sheets[0], 'A1', makeCell(5))]);
  assert.deepEqual(column(workbook, 'B1', 6), [1, 2, 3, 4, 5, undefined]);
  assert.deepEqual([shown(workbook, 'C1'), shown(workbook, 'D1'), shown(workbook, 'E1')], [15, 5, 5]);
  engine.recalculate([place(workbook.sheets[0], 'A1', makeCell(2))]);
  assert.deepEqual(column(workbook, 'B1', 6), [1, 2, undefined, undefined, undefined, undefined]);
  assert.deepEqual([shown(workbook, 'C1'), shown(workbook, 'D1'), shown(workbook, 'E1')], [3, 0, 2]);
  assert.equal((getCell(workbook.sheets[0], 0, 1) as Cell).formula?.ref, 'B1:B2');
});

test('the order formulas are calculated in does not matter: a reader done before the spill grew is done again', () => {
  // D1 is read first (it comes first in the list of what changed) but reads a cell B1 spills into.
  const workbook = book({ A1: 1, B1: '=SEQUENCE(A1)', D1: '=B3*10' });
  const engine = new CalcEngine(workbook);
  engine.recalculate(engine.uncalculated());
  assert.equal(shown(workbook, 'D1'), 0);
  const changed = place(workbook.sheets[0], 'A1', makeCell(4));
  engine.recalculate([{ sheet: 0, row: 0, col: 3 }, changed]);
  assert.equal(shown(workbook, 'D1'), 30);
});

test('a reader calculated before five spills grew into what it reads is calculated again for every one of them', () => {
  const workbook = book({ G1: 1, A1: '=SEQUENCE(G1)', B1: '=SEQUENCE(G1)', C1: '=SEQUENCE(G1)', D1: '=SEQUENCE(G1)', E1: '=SEQUENCE(G1)', H1: '=SUM(A2:E10)' });
  const engine = new CalcEngine(workbook);
  engine.recalculate(engine.uncalculated());
  assert.equal(shown(workbook, 'H1'), 0);
  // H1 comes first: it is done before any of the spills reach its range.
  const changed = place(workbook.sheets[0], 'G1', makeCell(3));
  engine.recalculate([{ sheet: 0, row: 0, col: 7 }, changed]);
  assert.equal(shown(workbook, 'H1'), 25);
  // As an engine that reads the workbook afresh finds it.
  const again = new CalcEngine(workbook);
  again.recalculate(undefined, { overwrite: true });
  assert.equal(shown(workbook, 'H1'), 25);
});

test('removing the formula empties the cells it spilled into', () => {
  const workbook = book({ A1: '=SEQUENCE(3)', C1: '=SUM(A1:A3)' });
  const sheet = workbook.sheets[0];
  const before = new Map(sheet.rows);
  const engine = new CalcEngine(workbook);
  engine.recalculate(engine.uncalculated());
  assert.equal(shown(workbook, 'C1'), 6);
  engine.recalculate([place(sheet, 'A1', undefined)]);
  assert.deepEqual(column(workbook, 'A1', 3), [undefined, undefined, undefined]);
  assert.equal(shown(workbook, 'C1'), 0);
  // The rows the spill made are gone again.
  assert.deepEqual(Array.from(sheet.rows.keys()).sort(), Array.from(before.keys()).sort());
});

test('a format put on a cell of the spill stays when the spill goes', () => {
  const workbook = book({ A1: '=SEQUENCE(3)', A2: { f: '', v: null } });
  const sheet = workbook.sheets[0];
  place(sheet, 'A2', { kind: 'blank', value: null, style: 4 });
  const engine = new CalcEngine(workbook);
  engine.recalculate(engine.uncalculated());
  assert.equal((getCell(sheet, 1, 0) as Cell).style, 4);
  assert.equal(shown(workbook, 'A2'), 2);
  engine.recalculate([place(sheet, 'A1', undefined)]);
  assert.deepEqual(getCell(sheet, 1, 0), { kind: 'blank', value: null, style: 4 });
});

test('two spills in each other\'s way: the second one is #SPILL!', () => {
  const { workbook } = calculate({ A1: '=SEQUENCE(1,3)', C1: 'x', A2: '=SEQUENCE(1,2)', B3: '=SEQUENCE(2)' });
  assert.equal(shown(workbook, 'A1'), '#SPILL!');
  assert.deepEqual([shown(workbook, 'A2'), shown(workbook, 'B2')], [1, 2]);
  assert.deepEqual(column(workbook, 'B3', 2), [1, 2]);
});

test('"#" is the whole spill; a cell that spills nothing gives #REF!', () => {
  const { workbook } = calculate({ A1: '=SEQUENCE(4)', B1: '=SUM(A1#)', C1: '=ROWS(A1#)', D1: 5, E1: '=D1#', F1: '=A1#*10' });
  assert.equal(shown(workbook, 'B1'), 10);
  assert.equal(shown(workbook, 'C1'), 4);
  assert.equal(shown(workbook, 'E1'), '#REF!');
  assert.deepEqual(column(workbook, 'F1', 4), [10, 20, 30, 40]);
});

test('a formula of one value that means the same as an ordinary one is written as one', () => {
  const { workbook } = calculate({ A1: 1, A2: 2, B1: '=A1+1', B2: '=SUM(A1:A2)', B3: '=A1:A2', C2: '=A1:A2+0' });
  const ref = (address: string) => {
    const at = parseCellRef(address) as { row: number; col: number };
    return (getCell(workbook.sheets[0], at.row, at.col) as Cell).formula?.ref;
  };
  assert.equal(ref('B1'), undefined);
  assert.equal(ref('B2'), undefined);
  // A range where one value is wanted: an ordinary formula would take one cell of it.
  assert.equal(ref('B3'), 'B3:B4');
  assert.equal(ref('C2'), 'C2:C3');
  // INDEX without a column may give a row: taken for a dynamic formula, of one value here.
  const single = calculate({ A1: 1, A2: 2, B1: '=INDEX(A1:A2*2,2)', C1: '=ABS(A1:A2)', D1: '=INDEX(A1:A2,2,1)' }).workbook;
  assert.equal((getCell(single.sheets[0], 0, 1) as Cell).formula?.ref, 'B1');
  assert.equal((getCell(single.sheets[0], 0, 3) as Cell).formula?.dynamic, undefined);
  assert.equal((getCell(single.sheets[0], 0, 2) as Cell).formula?.ref, 'C1:C2');
});

test('the values a file has for a dynamic array formula stay until something it reads changes', () => {
  const workbook = book({
    A1: 3,
    B1: { f: '_xlfn.SEQUENCE(A1)', dynamic: true, ref: 'B1:B3', v: 1 },
    B2: 2,
    B3: 3,
  });
  // As the reader marks them: the cells of the spill the file has.
  for (const address of ['B2', 'B3']) {
    const at = parseCellRef(address) as { row: number; col: number };
    place(workbook.sheets[0], address, { ...(getCell(workbook.sheets[0], at.row, at.col) as Cell), spilled: true });
  }
  const engine = new CalcEngine(workbook);
  const before = workbook.sheets[0].rows;
  engine.recalculate();
  assert.equal(workbook.sheets[0].rows, before);
  assert.equal(workbook.sheets[0].modified, undefined);
  engine.recalculate([place(workbook.sheets[0], 'A1', makeCell(4))]);
  assert.deepEqual(column(workbook, 'B1', 4), [1, 2, 3, 4]);
});

test('a formula the file saved as circular through its spill fills nothing, and takes nothing of the budget of spills (seed 7009)', () => {
  const make = () => {
    const workbook = book({
      // Spilling into B1, A1 would read its own spill: a circle, saved as 0 in its own cell.
      A1: { f: 'B1:C1', dynamic: true, ref: 'A1', v: 0 },
      A5: 3,
      E1: { f: '_xlfn.SEQUENCE(A5)', dynamic: true, ref: 'E1:E3', v: 1 },
      E2: 2,
      E3: 3,
    });
    for (const address of ['E2', 'E3']) {
      const at = parseCellRef(address) as { row: number; col: number };
      place(workbook.sheets[0], address, { ...(getCell(workbook.sheets[0], at.row, at.col) as Cell), spilled: true });
    }
    return workbook;
  };
  const options = { limits: { spillCells: 6, workbookSpillCells: 3 } };
  // The session opened the file, then A5 changed: E1 wants two cells, A1 none.
  const session = make();
  const engine = new CalcEngine(session, options);
  engine.recalculate(engine.uncalculated());
  engine.recalculate([place(session.sheets[0], 'A5', makeCell(2))]);
  // A workbook built afresh from the same cells.
  const fresh = make();
  place(fresh.sheets[0], 'A5', makeCell(2));
  new CalcEngine(fresh, options).recalculate(undefined, { overwrite: true });
  assert.deepEqual(column(fresh, 'E1', 3), [1, 2, undefined]);
  assert.deepEqual(column(session, 'E1', 3), [1, 2, undefined]);
  assert.equal(shown(session, 'A1'), 0);
  assert.equal(shown(session, 'B1'), undefined);
});

test('a volatile array is calculated by every participant alike when the random numbers are', () => {
  const make = () => {
    let seed = 7;
    const random = () => {
      seed = (seed * 16807) % 2147483647;
      return seed / 2147483647;
    };
    const workbook = book({ A1: '=SORT(RANDARRAY(5))', B1: '=SEQUENCE(3)*A1' });
    const engine = new CalcEngine(workbook, { random });
    engine.recalculate(engine.uncalculated());
    return workbook;
  };
  const one = make();
  const two = make();
  assert.deepEqual(column(one, 'A1', 5), column(two, 'A1', 5));
  assert.deepEqual(column(one, 'B1', 3), column(two, 'B1', 3));
});
