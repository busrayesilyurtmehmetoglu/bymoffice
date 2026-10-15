// Syntax tree -> value.
//
// Operators work on scalars and, element by element, on arrays; a reference stays a
// reference until something needs what is in its cells. What a range of several cells
// stands for then depends on the formula: in an array formula (and inside an argument that
// takes an array, such as those of SUMPRODUCT) it is the array of its cells; in an ordinary
// formula it is the one cell in the formula's row or column (implicit intersection). The functions that choose what to
// calculate (IF, IFERROR, CHOOSE, ...) or that need to know where cells are (ROW, OFFSET,
// INDEX, ...) are handled here, before the function library is asked.
//
// Nothing here throws for a bad formula: what cannot be calculated is an error value.

import { withoutFuturePrefixes } from '../edit/future-names';
import { MAX_COLS, MAX_ROWS, columnName, parseCellRef, type CellRange } from '../model/sheet/types';
import { APPLY_CALL, parseFormula, type BinaryOp, type Node, type RefNode } from './parse';
import {
  ERROR_DIV0,
  ERROR_NA,
  ERROR_NAME,
  ERROR_NUM,
  ERROR_REF,
  ERROR_VALUE,
  arrayAt,
  beginTextBudget,
  calcError,
  chargeText,
  checkNumber,
  compareScalars,
  endTextBudget,
  isArray,
  isError,
  isRef,
  refToArray,
  toBoolean,
  toNumber,
  toText,
  type ArgSelector,
  type ArrayValue,
  type CalcContext,
  type CalcError,
  type FunctionSpec,
  type RefValue,
  type Scalar,
  type Value,
} from './values';

export interface EvaluateOptions {
  /** The formula of a defined name as seen from a sheet, or undefined when there is no such name. */
  resolveName?: (name: string, sheet: number) => Node | undefined;
  /** The formula text of a cell without the leading "=", or undefined when the cell holds no formula. */
  cellFormula?: (sheet: number, row: number, col: number) => string | undefined;
  /**
   * Rows and columns to move the relative references of the tree by before they are used.
   * This is how the cells of a shared formula are calculated from the tree of their master.
   */
  rowShift?: number;
  colShift?: number;
  /**
   * The formula is an array formula (entered with Ctrl+Shift+Enter, or a dynamic array
   * formula): ranges are arrays everywhere, there is no implicit intersection.
   */
  arrayFormula?: boolean;
  /** The formula is a dynamic array formula: as `arrayFormula`, no implicit intersection ("@" asks for it). */
  dynamic?: boolean;
}

/**
 * A function made by LAMBDA: its parameters (in upper case), its body, and the names of LET
 * and of outer LAMBDAs it was made among. It is a value only inside a formula: a cell, an
 * operator or a function of the library that gets one gets #CALC!.
 */
export interface LambdaValue {
  readonly lambda: true;
  readonly params: readonly string[];
  readonly body: Node;
  readonly scope: Scope | undefined;
  readonly rowShift: number;
  readonly colShift: number;
}

/** A name of LET or a parameter of LAMBDA, with the names it was made among. */
interface Scope {
  readonly name: string;
  readonly value: Bound;
  readonly parent: Scope | undefined;
}

/** A parameter of LAMBDA that the call gave no argument for (see ISOMITTED). */
const OMITTED = { omitted: true } as const;

type Bound = Value | LambdaValue | typeof OMITTED;

const isOmitted = (value: Bound): value is typeof OMITTED => value === OMITTED;

function isLambda(value: unknown): value is LambdaValue {
  return typeof value === 'object' && value !== null && (value as LambdaValue).lambda === true;
}

/**
 * Something thrown to give up a calculation from inside (the engine does so to restart with
 * a shallower stack). It passes through the evaluator untouched; every other exception a
 * function throws becomes #VALUE!.
 */
export class EvaluationAbort {}

/** Nested calls, operators and names in one evaluate() call; deeper trees give #VALUE!. */
const MAX_NODE_DEPTH = 600;

/** Largest array an operator or a function here builds. */
const MAX_ARRAY_CELLS = 5000000;

/** Excel's limit on the length of a text value. */
const MAX_TEXT_LENGTH = 32767;

/**
 * Calls of LAMBDA inside one another (a name that calls itself): deeper ones give #NUM!.
 * UNSURE: Excel's own limit and its error were not checked against Excel.
 */
export const MAX_LAMBDA_DEPTH = 100;

/** Parameters of one LAMBDA and names of one LET; Excel allows 253 of either. */
const MAX_PARAMS = 253;

const ERROR_NULL = calcError('#NULL!');
const ERROR_CALC = calcError('#CALC!');

interface State {
  context: CalcContext;
  functions: Map<string, FunctionSpec>;
  options: EvaluateOptions;
  /** Value of `depth` when this evaluation started. */
  base: number;
  rowShift: number;
  colShift: number;
  /** Names being worked out, to stop a name that refers to itself. */
  names: string[];
  /** Ranges are arrays here: the formula is an array formula, or this is an argument that takes an array. */
  arrayMode: boolean;
  /** What `arrayMode` would be were the formula an ordinary one: false in a dynamic array formula but inside arguments that take arrays. */
  legacyArray: boolean;
  /** The formula is a dynamic array formula (see EvaluateOptions.dynamic). */
  dynamic: boolean;
  /** The names of LET and the parameters of LAMBDA that are seen here. */
  scope: Scope | undefined;
  /** Calls of LAMBDA in progress, one inside the other. */
  lambdas: number;
}

// How deep the evaluator is, counted over every evaluation in progress: a cell that needs
// another cell starts a new evaluate() call inside the first one.
let depth = 0;

/** How many nested operations all evaluations in progress add up to; a measure of stack use. */
export function evaluationDepth(): number {
  return depth;
}

/**
 * Calculates a tree. References come back as references and arrays as arrays; use
 * toCellValue() for what a single cell would show.
 */
export function evaluate(node: Node, context: CalcContext, functions: Map<string, FunctionSpec>, options?: EvaluateOptions): Value {
  const opts = options || {};
  const dynamic = opts.dynamic === true;
  const state: State = {
    context,
    functions,
    options: opts,
    base: depth,
    rowShift: opts.rowShift || 0,
    colShift: opts.colShift || 0,
    names: [],
    arrayMode: opts.arrayFormula === true || dynamic,
    legacyArray: opts.arrayFormula === true && !dynamic,
    dynamic,
    scope: undefined,
    lambdas: 0,
  };
  const text = beginTextBudget();
  try {
    return evalNode(node, state);
  } finally {
    depth = state.base;
    endTextBudget(text);
  }
}

/**
 * Implicit intersection: the one cell a reference stands for in an ordinary formula. A single
 * cell is itself; of one column it is the cell in the formula's row, of one row the cell in
 * its column; #VALUE! when there is no such cell.
 */
export function intersect(ref: RefValue, context: CalcContext): Scalar {
  const range = ref.range;
  if (range.top === range.bottom && range.left === range.right) return context.cellValue(ref.sheet, range.top, range.left);
  if (range.left === range.right && context.row >= range.top && context.row <= range.bottom) return context.cellValue(ref.sheet, context.row, range.left);
  if (range.top === range.bottom && context.col >= range.left && context.col <= range.right) return context.cellValue(ref.sheet, range.top, context.col);
  return ERROR_VALUE;
}

/** Result as it is stored in one cell: a reference is intersected, of an array the first element is taken. */
export function toCellValue(value: Value, context: CalcContext): Scalar {
  if (isRef(value)) return intersect(value, context);
  if (isArray(value)) return value.data.length > 0 ? value.data[0] : ERROR_VALUE;
  return value;
}

/**
 * The cells a reference node points at once its relative sides are moved by (rowShift,
 * colShift); undefined when that pushes it off the sheet.
 */
export function refRange(node: RefNode, rowShift: number, colShift: number): CellRange | undefined {
  let { top, left, bottom, right } = node.range;
  if (rowShift === 0 && colShift === 0) return node.range;
  if (!node.wholeCols) {
    if (!node.abs.top) top += rowShift;
    if (!node.abs.bottom) bottom += rowShift;
    if (top < 0 || bottom < 0 || top >= MAX_ROWS || bottom >= MAX_ROWS) return undefined;
  }
  if (!node.wholeRows) {
    if (!node.abs.left) left += colShift;
    if (!node.abs.right) right += colShift;
    if (left < 0 || right < 0 || left >= MAX_COLS || right >= MAX_COLS) return undefined;
  }
  return { top: Math.min(top, bottom), left: Math.min(left, right), bottom: Math.max(top, bottom), right: Math.max(left, right) };
}

// ---------------------------------------------------------------------------
// Nodes

function evalNode(node: Node, state: State): Value {
  switch (node.type) {
    case 'number':
    case 'string':
    case 'boolean':
      return node.value;
    case 'error':
      return calcError(node.code);
    case 'ref':
      return evalRef(node, state);
    case 'array':
      return evalArrayConstant(node.rows);
  }
  if (depth - state.base >= MAX_NODE_DEPTH) return ERROR_VALUE;
  depth++;
  const result = node.type === 'binary' ? evalBinary(node, state) : node.type === 'call' ? evalCall(node, state) : node.type === 'unary' ? evalUnary(node, state) : evalName(node, state);
  depth--;
  return result;
}

function evalRef(node: RefNode, state: State): RefValue | CalcError {
  // References into other workbooks and across sheets (3D) are not supported.
  if (node.external || node.sheetEnd !== undefined) return ERROR_REF;
  let sheet = state.context.sheet;
  if (node.sheet !== undefined) {
    const index = state.context.sheetIndex(node.sheet);
    if (index === undefined) return ERROR_REF;
    sheet = index;
  }
  const range = state.rowShift === 0 && state.colShift === 0 ? node.range : refRange(node, state.rowShift, state.colShift);
  return range ? { ref: true, sheet, range } : ERROR_REF;
}

function evalArrayConstant(rows: Node[][]): ArrayValue {
  const cols = rows.length > 0 ? rows[0].length : 0;
  const data: Scalar[] = [];
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < cols; c++) {
      const item = rows[r][c];
      data.push(!item ? null : item.type === 'error' ? calcError(item.code) : item.type === 'number' || item.type === 'string' || item.type === 'boolean' ? item.value : ERROR_VALUE);
    }
  }
  return { rows: rows.length, cols, data };
}

/** What a name of LET or a parameter of LAMBDA is bound to, or undefined when no such name is seen here. */
function scoped(name: string, state: State): { value: Bound } | undefined {
  if (!state.scope) return undefined;
  const upper = name.toUpperCase();
  for (let scope: Scope | undefined = state.scope; scope; scope = scope.parent) if (scope.name === upper) return scope;
  return undefined;
}

function evalName(node: Extract<Node, { type: 'name' }>, state: State): Value {
  if (node.sheet === undefined) {
    const local = scoped(node.name, state);
    if (local) return isOmitted(local.value) ? null : isLambda(local.value) ? ERROR_CALC : local.value;
  }
  const resolve = state.options.resolveName;
  let sheet = state.context.sheet;
  if (node.sheet !== undefined) {
    const index = state.context.sheetIndex(node.sheet);
    if (index === undefined) return ERROR_REF;
    sheet = index;
  }
  if (!resolve) return ERROR_NAME;
  const key = `${sheet}!${node.name.toUpperCase()}`;
  // A name that leads back to itself has no value.
  if (state.names.indexOf(key) >= 0) return ERROR_NAME;
  const target = resolve(node.name, sheet);
  if (!target) return ERROR_NAME;
  const rowShift = state.rowShift;
  const colShift = state.colShift;
  state.names.push(key);
  state.rowShift = 0;
  state.colShift = 0;
  const result = evalNode(target, state);
  state.rowShift = rowShift;
  state.colShift = colShift;
  state.names.pop();
  return result;
}

/** What is in the cells of a reference: a scalar for one cell, an array for more. */
function derefAll(value: Value, state: State): Scalar | ArrayValue {
  if (!isRef(value)) return value;
  const range = value.range;
  if (range.top === range.bottom && range.left === range.right) return state.context.cellValue(value.sheet, range.top, range.left);
  return refToArray(value, state.context);
}

/** What a reference stands for where a value is needed: its cells as an array, or in an ordinary formula the one cell it intersects. */
function deref(value: Value, state: State): Scalar | ArrayValue {
  if (!isRef(value)) return value;
  return state.arrayMode ? derefAll(value, state) : intersect(value, state.context);
}

const isSingleCell = (ref: RefValue) => ref.range.top === ref.range.bottom && ref.range.left === ref.range.right;

/** Calculates an argument that takes an array: ranges inside it are arrays whatever the formula is. */
function evalArrayArg(arg: Node, state: State): Value {
  if (state.legacyArray) return evalNode(arg, state);
  const mode = state.arrayMode;
  state.arrayMode = true;
  state.legacyArray = true;
  const value = evalNode(arg, state);
  state.arrayMode = mode;
  state.legacyArray = false;
  return value;
}

const selects = (selector: ArgSelector | undefined, index: number): boolean => selector === true || (selector !== undefined && selector(index) === true);

// ---------------------------------------------------------------------------
// Operators

function power(x: number, y: number): Scalar {
  if (x === 0) {
    if (y === 0) return ERROR_NUM;
    if (y < 0) return ERROR_DIV0;
  }
  if (x < 0 && Math.floor(y) !== y) return ERROR_NUM;
  return checkNumber(Math.pow(x, y));
}

function scalarBinary(op: BinaryOp, a: Scalar, b: Scalar, context: CalcContext): Scalar {
  switch (op) {
    case '+':
    case '-':
    case '*':
    case '/':
    case '^': {
      const x = typeof a === 'number' ? a : toNumber(a, context);
      if (typeof x !== 'number') return x;
      const y = typeof b === 'number' ? b : toNumber(b, context);
      if (typeof y !== 'number') return y;
      if (op === '+') return checkNumber(x + y);
      if (op === '-') return checkNumber(x - y);
      if (op === '*') return checkNumber(x * y);
      if (op === '/') return y === 0 ? ERROR_DIV0 : checkNumber(x / y);
      return power(x, y);
    }
    case '&': {
      const x = toText(a);
      if (typeof x !== 'string') return x;
      const y = toText(b);
      if (typeof y !== 'string') return y;
      return x.length + y.length > MAX_TEXT_LENGTH ? ERROR_VALUE : x + y;
    }
    default: {
      const order = compareScalars(a, b);
      if (typeof order !== 'number') return order;
      switch (op) {
        case '=':
          return order === 0;
        case '<>':
          return order !== 0;
        case '<':
          return order < 0;
        case '>':
          return order > 0;
        case '<=':
          return order <= 0;
        case '>=':
          return order >= 0;
      }
      return ERROR_VALUE;
    }
  }
}

/**
 * Applies `fn` to every pair of elements; a scalar or a single row or column stands for every position.
 * #VALUE! when the result would be too large, in cells or in the text it holds.
 */
function zip(a: Scalar | ArrayValue, b: Scalar | ArrayValue, state: State, fn: (x: Scalar, y: Scalar) => Scalar): Scalar | ArrayValue {
  const left = isArray(a) ? a : undefined;
  const right = isArray(b) ? b : undefined;
  if (!left && !right) return fn(a as Scalar, b as Scalar);
  const rows = Math.max(left ? left.rows : 1, right ? right.rows : 1);
  const cols = Math.max(left ? left.cols : 1, right ? right.cols : 1);
  if (!(rows * cols <= MAX_ARRAY_CELLS)) return ERROR_VALUE;
  const fast = zipTail(left, right, a, b, rows, cols, state, fn);
  if (fast) return fast;
  if (state.context.tick) state.context.tick(rows * cols);
  const data: Scalar[] = new Array(rows * cols);
  let k = 0;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const value = fn(left ? arrayAt(left, r, c) : (a as Scalar), right ? arrayAt(right, r, c) : (b as Scalar));
      if (typeof value === 'string' && !chargeText(value)) return ERROR_VALUE;
      data[k++] = value;
    }
  }
  return { rows, cols, data };
}

/**
 * zip() for arrays of the result's shape (or scalars) with tails (ArrayValue.tail): the elements before the
 * tails are worked out one by one, those after from one pair; the same result, without a call per element
 * of the empty part of a whole column. Undefined when it does not apply.
 */
function zipTail(left: ArrayValue | undefined, right: ArrayValue | undefined, a: Scalar | ArrayValue, b: Scalar | ArrayValue, rows: number, cols: number, state: State, fn: (x: Scalar, y: Scalar) => Scalar): ArrayValue | undefined {
  const size = rows * cols;
  const fits = (array: ArrayValue | undefined) => !array || (array.rows === rows && array.cols === cols && array.tail !== undefined && array.data.length >= size);
  if (!fits(left) || !fits(right) || (!left && !right)) return undefined;
  const from = Math.max(left ? (left.tail as number) : 0, right ? (right.tail as number) : 0);
  if (!(from < size)) return undefined;
  const rest = fn(left ? left.data[from] : (a as Scalar), right ? right.data[from] : (b as Scalar));
  // Text is charged by the element (chargeText()): not worked out this way.
  if (typeof rest === 'string') return undefined;
  if (state.context.tick) state.context.tick(from + (size >> 4));
  const data: Scalar[] = new Array(size);
  for (let k = 0; k < from; k++) {
    const value = fn(left ? left.data[k] : (a as Scalar), right ? right.data[k] : (b as Scalar));
    if (typeof value === 'string' && !chargeText(value)) return { rows: 1, cols: 1, data: [ERROR_VALUE] };
    data[k] = value;
  }
  data.fill(rest, from);
  return { rows, cols, data, tail: from };
}

function mapValues(a: Scalar | ArrayValue, state: State, fn: (x: Scalar) => Scalar): Scalar | ArrayValue {
  if (!isArray(a)) return fn(a);
  const tail = a.tail;
  if (tail !== undefined && tail < a.data.length) {
    const rest = fn(a.data[tail]);
    if (typeof rest !== 'string') {
      if (state.context.tick) state.context.tick(tail + (a.data.length >> 4));
      const data: Scalar[] = new Array(a.data.length);
      for (let k = 0; k < tail; k++) data[k] = fn(a.data[k]);
      data.fill(rest, tail);
      return { rows: a.rows, cols: a.cols, data, tail };
    }
  }
  if (state.context.tick) state.context.tick(a.data.length);
  const data: Scalar[] = new Array(a.data.length);
  for (let k = 0; k < data.length; k++) data[k] = fn(a.data[k]);
  return { rows: a.rows, cols: a.cols, data };
}

function refOperator(op: BinaryOp, a: Value, b: Value): Value {
  if (isError(a)) return a;
  if (isError(b)) return b;
  if (!isRef(a) || !isRef(b)) return ERROR_VALUE;
  // A union is only understood where a function takes it apart (see evalCall).
  if (op === ',') return ERROR_VALUE;
  const x = a.range;
  const y = b.range;
  if (op === ':') {
    if (a.sheet !== b.sheet) return ERROR_REF;
    return { ref: true, sheet: a.sheet, range: { top: Math.min(x.top, y.top), left: Math.min(x.left, y.left), bottom: Math.max(x.bottom, y.bottom), right: Math.max(x.right, y.right) } };
  }
  const range = { top: Math.max(x.top, y.top), left: Math.max(x.left, y.left), bottom: Math.min(x.bottom, y.bottom), right: Math.min(x.right, y.right) };
  if (a.sheet !== b.sheet || range.top > range.bottom || range.left > range.right) return ERROR_NULL;
  return { ref: true, sheet: a.sheet, range };
}

const isRefOp = (op: BinaryOp) => op === ':' || op === ' ' || op === ',';

type BinaryNode = Extract<Node, { type: 'binary' }>;

// Operators group from the left, so a long formula such as 1+1+1+... is a tree that is deep
// on its left side only. That side is walked in a loop rather than by recursion.
function evalBinary(node: BinaryNode, state: State): Value {
  const spine: BinaryNode[] = [];
  let first: Node = node;
  while (first.type === 'binary') {
    spine.push(first);
    first = first.left;
  }
  let result = evalNode(first, state);
  for (let k = spine.length - 1; k >= 0; k--) {
    const op = spine[k].op;
    const right = evalNode(spine[k].right, state);
    if (isRefOp(op)) {
      result = refOperator(op, result, right);
      continue;
    }
    const a = deref(result, state);
    const b = deref(right, state);
    const context = state.context;
    result = isArray(a) || isArray(b) ? zip(a, b, state, (x, y) => scalarBinary(op, x, y, context)) : scalarBinary(op, a, b, context);
  }
  return result;
}

function scalarUnary(op: '-' | '+' | '%', value: Scalar, context: CalcContext): Scalar {
  // A plus sign changes nothing, not even text.
  if (op === '+') return value;
  const x = toNumber(value, context);
  if (typeof x !== 'number') return x;
  if (op === '%') return x / 100;
  return x === 0 ? 0 : -x;
}

function evalUnary(node: Extract<Node, { type: 'unary' }>, state: State): Value {
  const ops: ('-' | '+' | '%')[] = [];
  let operand: Node = node;
  while (operand.type === 'unary') {
    ops.push(operand.op);
    operand = operand.operand;
  }
  let result = deref(evalNode(operand, state), state);
  for (let k = ops.length - 1; k >= 0; k--) {
    const op = ops[k];
    result = mapValues(result, state, (x) => scalarUnary(op, x, state.context));
  }
  return result;
}

// ---------------------------------------------------------------------------
// Function calls

type Args = (Node | null)[];
type Special = (args: Args, state: State) => Value;

/**
 * The references a union is made of, in the order they are written. A union inside a
 * union, on either side ("(A1:A3,(B1:B3,C1:C3))"), is taken apart as well.
 */
export function unionParts(node: Node): Node[] {
  const parts: Node[] = [];
  // What is still to be looked at, the leftmost on top; no deeper than the tree, which the parser bounds.
  const pending: Node[] = [node];
  while (pending.length > 0) {
    const current = pending.pop() as Node;
    if (current.type === 'binary' && current.op === ',') pending.push(current.right, current.left);
    else parts.push(current);
  }
  return parts;
}

function evalCall(node: Extract<Node, { type: 'call' }>, state: State): Value {
  const special = SPECIALS.get(node.name);
  if (special) return special(node.args, state);
  const spec = state.functions.get(node.name);
  if (!spec) {
    // A function the formula makes itself: a name of LET or a defined name that stands for a LAMBDA.
    const made = calledLambda(node.name, state);
    if (made === undefined) return ERROR_NAME;
    if (!isLambda(made)) return made;
    return valueOf(invoke(made, node.args, state));
  }
  const nodes = node.args;
  // Excel refuses to accept such a formula; here it is an error value.
  if (nodes.length < spec.min || (spec.max !== undefined && nodes.length > spec.max)) return ERROR_VALUE;
  const args: Value[] = [];
  const ordinary = !state.arrayMode;
  for (let k = 0; k < nodes.length; k++) {
    const arg = nodes[k];
    if (arg === null) args.push(null);
    else if (arg.type === 'binary' && arg.op === ',' && spec.max === undefined) {
      // SUM((A1:A3,C1:C3)): a function that takes any number of arguments gets the parts of a union one by one.
      const parts = unionParts(arg);
      for (let j = 0; j < parts.length; j++) args.push(evalNode(parts[j], state));
    } else if (!state.legacyArray && selects(spec.arrayArgs, k)) args.push(evalArrayArg(arg, state));
    else {
      const value = evalNode(arg, state);
      // Where the function takes one value, a range stands for the cell it intersects.
      args.push(ordinary && isRef(value) && selects(spec.scalarArgs, k) ? intersect(value, state.context) : value);
    }
  }
  if (state.context.tick) state.context.tick(1);
  const saved = depth;
  let result: Value;
  try {
    result = spec.call(args, state.context);
  } catch (error) {
    if (error instanceof EvaluationAbort) throw error;
    depth = saved;
    return ERROR_VALUE;
  }
  if (typeof result === 'number') return checkNumber(result);
  if (result === undefined || (typeof result !== 'object' && typeof result !== 'string' && typeof result !== 'boolean')) return ERROR_VALUE;
  return result;
}

const evalArg = (arg: Node | null | undefined, state: State): Value => (arg ? evalNode(arg, state) : null);

const scalarArg = (arg: Node | null | undefined, state: State): Scalar => (arg ? toCellValue(evalNode(arg, state), state.context) : null);

/** An argument as a whole number, cut towards zero; an empty one is 0. */
function intArg(arg: Node | null | undefined, state: State): number | CalcError {
  const number = toNumber(scalarArg(arg, state), state.context);
  return typeof number === 'number' ? Math.trunc(number) : number;
}

/** The first cell of a reference or an array, or the scalar itself. */
function firstValue(value: Value, state: State): Scalar {
  if (isRef(value)) return state.context.cellValue(value.sheet, value.range.top, value.range.left);
  if (isArray(value)) return value.data.length > 0 ? value.data[0] : null;
  return value;
}

/** A branch of IF and its relatives; an argument left empty counts as 0. */
const branch = (arg: Node | null | undefined, state: State): Value => (arg ? evalNode(arg, state) : 0);

function fnIf(args: Args, state: State): Value {
  if (args.length < 2 || args.length > 3) return ERROR_VALUE;
  const condition = deref(evalArg(args[0], state), state);
  if (!isArray(condition)) {
    const test = toBoolean(condition);
    if (typeof test !== 'boolean') return test;
    if (test) return branch(args[1], state);
    return args.length < 3 ? false : branch(args[2], state);
  }
  // An array of conditions chooses element by element; a branch nothing chooses is still not calculated.
  let anyTrue = false;
  let anyFalse = false;
  const tests: (boolean | CalcError)[] = condition.data.map((value) => {
    const test = toBoolean(value);
    if (test === true) anyTrue = true;
    else if (test === false) anyFalse = true;
    return test;
  });
  const whenTrue = anyTrue ? derefAll(branch(args[1], state), state) : null;
  const whenFalse = !anyFalse ? null : args.length < 3 ? false : derefAll(branch(args[2], state), state);
  const yes = isArray(whenTrue) ? whenTrue : undefined;
  const no = isArray(whenFalse) ? whenFalse : undefined;
  const rows = Math.max(condition.rows, yes ? yes.rows : 1, no ? no.rows : 1);
  const cols = Math.max(condition.cols, yes ? yes.cols : 1, no ? no.cols : 1);
  if (!(rows * cols <= MAX_ARRAY_CELLS)) return ERROR_VALUE;
  if (state.context.tick) state.context.tick(rows * cols);
  const flags: ArrayValue = { rows: condition.rows, cols: condition.cols, data: tests };
  const data: Scalar[] = new Array(rows * cols);
  let k = 0;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const test = arrayAt(flags, r, c);
      if (test === true) data[k++] = yes ? arrayAt(yes, r, c) : (whenTrue as Scalar);
      else if (test === false) data[k++] = no ? arrayAt(no, r, c) : (whenFalse as Scalar);
      else data[k++] = test;
    }
  }
  return { rows, cols, data };
}

function ifCaught(args: Args, state: State, caught: (value: Scalar) => boolean): Value {
  if (args.length !== 2) return ERROR_VALUE;
  const value = deref(evalArg(args[0], state), state);
  if (!isArray(value)) return caught(value) ? evalArg(args[1], state) : value;
  if (!value.data.some(caught)) return value;
  const fallback = derefAll(evalArg(args[1], state), state);
  return zip(value, fallback, state, (x, y) => (caught(x) ? y : x));
}

function fnIfs(args: Args, state: State): Value {
  if (args.length < 2 || args.length % 2 !== 0) return ERROR_VALUE;
  for (let k = 0; k < args.length; k += 2) {
    const test = toBoolean(scalarArg(args[k], state));
    if (typeof test !== 'boolean') return test;
    if (test) return branch(args[k + 1], state);
  }
  return ERROR_NA;
}

function fnSwitch(args: Args, state: State): Value {
  if (args.length < 3) return ERROR_VALUE;
  const subject = scalarArg(args[0], state);
  if (isError(subject)) return subject;
  let k = 1;
  for (; k + 1 < args.length; k += 2) {
    const candidate = scalarArg(args[k], state);
    if (isError(candidate)) return candidate;
    if (compareScalars(subject, candidate) === 0) return branch(args[k + 1], state);
  }
  // One argument left over is the default.
  return k < args.length ? branch(args[k], state) : ERROR_NA;
}

function fnChoose(args: Args, state: State): Value {
  if (args.length < 2) return ERROR_VALUE;
  const index = intArg(args[0], state);
  if (typeof index !== 'number') return index;
  if (index < 1 || index >= args.length) return ERROR_VALUE;
  return branch(args[index], state);
}

/** AND, OR and XOR: how many logical values there are and how many of them are true. */
function logical(args: Args, state: State, combine: (trues: number, count: number) => boolean): Value {
  if (args.length === 0) return ERROR_VALUE;
  let count = 0;
  let trues = 0;
  let error: CalcError | undefined;
  // Inside a reference or an array, text and empty cells are passed over.
  const visit = (value: Scalar) => {
    if (typeof value === 'boolean' || typeof value === 'number') {
      count++;
      if (value) trues++;
    } else if (isError(value) && !error) error = value;
  };
  for (let k = 0; k < args.length; k++) {
    const value = evalArg(args[k], state);
    if (isRef(value)) state.context.forEachCell(value.sheet, value.range, visit);
    else if (isArray(value)) value.data.forEach(visit);
    else {
      // Given directly, text is read ("TRUE") and an empty argument is FALSE.
      const test = toBoolean(value);
      if (typeof test !== 'boolean') {
        if (!error) error = test;
      } else {
        count++;
        if (test) trues++;
      }
    }
  }
  if (error) return error;
  if (count === 0) return ERROR_VALUE;
  return combine(trues, count);
}

function fnNot(args: Args, state: State): Value {
  if (args.length !== 1) return ERROR_VALUE;
  return mapValues(deref(evalArg(args[0], state), state), state, (value) => {
    const test = toBoolean(value);
    return typeof test === 'boolean' ? !test : test;
  });
}

/** The reference an argument evaluates to; an error when it is anything else. */
function refArg(arg: Node | null | undefined, state: State): RefValue | CalcError {
  const value = evalArg(arg, state);
  if (isRef(value)) return value;
  return isError(value) ? value : ERROR_VALUE;
}

function rowOrColumn(args: Args, state: State, rows: boolean): Value {
  if (args.length > 1) return ERROR_VALUE;
  if (args.length === 0 || args[0] === null) return (rows ? state.context.row : state.context.col) + 1;
  const ref = refArg(args[0], state);
  if (!isRef(ref)) return ref;
  const first = rows ? ref.range.top : ref.range.left;
  const count = (rows ? ref.range.bottom : ref.range.right) - first + 1;
  // Every row or column of a range only where an array is wanted; an ordinary formula gets the first one.
  if (count === 1 || !state.arrayMode) return first + 1;
  const data: Scalar[] = new Array(count);
  for (let k = 0; k < count; k++) data[k] = first + k + 1;
  return rows ? { rows: count, cols: 1, data } : { rows: 1, cols: count, data };
}

function rowsOrColumns(args: Args, state: State, rows: boolean): Value {
  if (args.length !== 1) return ERROR_VALUE;
  const value = evalArg(args[0], state);
  if (isError(value)) return value;
  if (isRef(value)) return rows ? value.range.bottom - value.range.top + 1 : value.range.right - value.range.left + 1;
  if (isArray(value)) return rows ? value.rows : value.cols;
  return 1;
}

function fnAreas(args: Args, state: State): Value {
  if (args.length !== 1 || !args[0]) return ERROR_VALUE;
  const parts = unionParts(args[0]);
  for (let k = 0; k < parts.length; k++) {
    const ref = refArg(parts[k], state);
    if (!isRef(ref)) return ref;
  }
  return parts.length;
}

function fnOffset(args: Args, state: State): Value {
  if (args.length < 3 || args.length > 5) return ERROR_VALUE;
  const base = refArg(args[0], state);
  if (!isRef(base)) return base;
  const rows = intArg(args[1], state);
  if (typeof rows !== 'number') return rows;
  const cols = intArg(args[2], state);
  if (typeof cols !== 'number') return cols;
  let height: number | CalcError = base.range.bottom - base.range.top + 1;
  let width: number | CalcError = base.range.right - base.range.left + 1;
  if (args.length > 3 && args[3]) height = intArg(args[3], state);
  if (typeof height !== 'number') return height;
  if (args.length > 4 && args[4]) width = intArg(args[4], state);
  if (typeof width !== 'number') return width;
  if (height === 0 || width === 0) return ERROR_REF;
  let top = base.range.top + rows;
  let left = base.range.left + cols;
  // A negative size counts back from the starting cell.
  if (height < 0) {
    top += height + 1;
    height = -height;
  }
  if (width < 0) {
    left += width + 1;
    width = -width;
  }
  const bottom = top + height - 1;
  const right = left + width - 1;
  if (!(top >= 0 && left >= 0 && bottom < MAX_ROWS && right < MAX_COLS)) return ERROR_REF;
  return { ref: true, sheet: base.sheet, range: { top, left, bottom, right } };
}

/** Whether a tree is nothing but a reference: cells, a name, or references joined by ":" or a space. */
function isReferenceTree(node: Node): boolean {
  let current = node;
  for (let guard = 0; guard < 64; guard++) {
    if (current.type === 'ref' || current.type === 'name') return true;
    if (current.type !== 'binary' || (current.op !== ':' && current.op !== ' ')) return false;
    if (current.right.type !== 'ref' && current.right.type !== 'name') return false;
    current = current.left;
  }
  return false;
}

function fnIndirect(args: Args, state: State): Value {
  if (args.length < 1 || args.length > 2) return ERROR_VALUE;
  const text = toText(scalarArg(args[0], state));
  if (typeof text !== 'string') return text;
  if (args.length > 1 && args[1]) {
    const a1 = toBoolean(scalarArg(args[1], state));
    if (typeof a1 !== 'boolean') return a1;
    // R1C1 notation is not supported.
    if (!a1) return ERROR_REF;
  }
  let node: Node;
  try {
    if (text.trim().charAt(0) === '=') return ERROR_REF;
    node = parseFormula(text);
  } catch (error) {
    return ERROR_REF;
  }
  if (!isReferenceTree(node)) return ERROR_REF;
  // The text states the cells as they are: nothing of a shared formula's shift applies to it.
  const rowShift = state.rowShift;
  const colShift = state.colShift;
  state.rowShift = 0;
  state.colShift = 0;
  const result = evalNode(node, state);
  state.rowShift = rowShift;
  state.colShift = colShift;
  return isRef(result) ? result : ERROR_REF;
}

function fnIndex(args: Args, state: State): Value {
  if (args.length < 2 || args.length > 4) return ERROR_VALUE;
  let area = 1;
  if (args.length > 3 && args[3]) {
    const given = intArg(args[3], state);
    if (typeof given !== 'number') return given;
    area = given;
  }
  // The reference form: the first argument may be several areas, "(A1:B2,D1:E2)", and the fourth picks one.
  const areas = args[0] ? unionParts(args[0]) : [];
  let source: Value;
  if (areas.length > 1) {
    for (let k = 0; k < areas.length; k++) {
      const ref = refArg(areas[k], state);
      if (!isRef(ref)) return ref;
    }
    if (area < 1) return ERROR_VALUE;
    if (area > areas.length) return ERROR_REF;
    source = refArg(areas[area - 1], state);
  } else {
    // The first argument takes an array: INDEX(A1:A3*2,3) needs no array formula.
    source = args[0] ? evalArrayArg(args[0], state) : null;
    if (isError(source)) return source;
    // There is one area only. UNSURE: Excel gives #VALUE! for an area below 1; taken from its documentation.
    if (area < 1) return ERROR_VALUE;
    if (area !== 1) return ERROR_REF;
  }
  let row = intArg(args[1], state);
  if (typeof row !== 'number') return row;
  let col: number | CalcError | undefined = args.length > 2 && args[2] ? intArg(args[2], state) : undefined;
  if (col !== undefined && typeof col !== 'number') return col;
  const array = isRef(source) ? undefined : isArray(source) ? source : { rows: 1, cols: 1, data: [source] };
  const rows = isRef(source) ? source.range.bottom - source.range.top + 1 : (array as ArrayValue).rows;
  const cols = isRef(source) ? source.range.right - source.range.left + 1 : (array as ArrayValue).cols;
  // With a single row, one index counts along it.
  if (col === undefined && rows === 1 && cols > 1) {
    col = row;
    row = 1;
  }
  if (col === undefined) col = cols === 1 ? 1 : 0;
  if (row < 0 || col < 0) return ERROR_VALUE;
  if (row > rows || col > cols) return ERROR_REF;
  // 0 stands for the whole row or column.
  if (isRef(source)) {
    const range = source.range;
    return {
      ref: true,
      sheet: source.sheet,
      range: {
        top: row ? range.top + row - 1 : range.top,
        left: col ? range.left + col - 1 : range.left,
        bottom: row ? range.top + row - 1 : range.bottom,
        right: col ? range.left + col - 1 : range.right,
      },
    };
  }
  const from = array as ArrayValue;
  if (row && col) return from.data[(row - 1) * from.cols + (col - 1)];
  if (!row && !col) return from;
  const data: Scalar[] = [];
  if (row) for (let c = 0; c < from.cols; c++) data.push(from.data[(row - 1) * from.cols + c]);
  else for (let r = 0; r < from.rows; r++) data.push(from.data[r * from.cols + (col - 1)]);
  return row ? { rows: 1, cols: from.cols, data } : { rows: from.rows, cols: 1, data };
}

function fnIsFormula(args: Args, state: State): Value {
  if (args.length !== 1) return ERROR_VALUE;
  const ref = refArg(args[0], state);
  if (!isRef(ref)) return ref;
  const lookup = state.options.cellFormula;
  return !!lookup && lookup(ref.sheet, ref.range.top, ref.range.left) !== undefined;
}

function fnFormulaText(args: Args, state: State): Value {
  if (args.length !== 1) return ERROR_VALUE;
  const ref = refArg(args[0], state);
  if (!isRef(ref)) return ref;
  const lookup = state.options.cellFormula;
  const text = lookup ? lookup(ref.sheet, ref.range.top, ref.range.left) : undefined;
  // As the formula is shown: without what files add to newer functions.
  return text === undefined ? ERROR_NA : `=${withoutFuturePrefixes(text)}`;
}

function fnSheet(args: Args, state: State): Value {
  if (args.length > 1) return ERROR_VALUE;
  if (args.length === 0 || !args[0]) return state.context.sheet + 1;
  const value = evalArg(args[0], state);
  if (isRef(value)) return value.sheet + 1;
  if (isError(value)) return value;
  if (typeof value === 'string') {
    const index = state.context.sheetIndex(value);
    return index === undefined ? ERROR_NA : index + 1;
  }
  return ERROR_NA;
}

function fnSheets(args: Args, state: State): Value {
  if (args.length > 1) return ERROR_VALUE;
  if (args.length === 0 || !args[0]) return state.context.sheetCount();
  const ref = refArg(args[0], state);
  return isRef(ref) ? 1 : ref;
}

function fnIsBlank(args: Args, state: State): Value {
  if (args.length !== 1) return ERROR_VALUE;
  return mapValues(deref(evalArg(args[0], state), state), state, (value) => value === null);
}

function fnN(args: Args, state: State): Value {
  if (args.length !== 1) return ERROR_VALUE;
  const value = firstValue(evalArg(args[0], state), state);
  if (typeof value === 'number' || isError(value)) return value;
  return value === true ? 1 : 0;
}

function fnT(args: Args, state: State): Value {
  if (args.length !== 1) return ERROR_VALUE;
  const value = firstValue(evalArg(args[0], state), state);
  return typeof value === 'string' || isError(value) ? value : '';
}

/** A sheet name as it is written in front of a reference. */
function quoteSheetName(name: string): string {
  let plain = name !== '' && parseCellRef(name) === undefined;
  for (let k = 0; plain && k < name.length; k++) {
    const c = name.charCodeAt(k);
    const letter = c >= 128 || (c >= 65 && c <= 90) || (c >= 97 && c <= 122) || c === 95;
    if (!letter && !(k > 0 && ((c >= 48 && c <= 57) || c === 46))) plain = false;
  }
  return plain ? name : `'${name.replace(/'/g, "''")}'`;
}

function fnCell(args: Args, state: State): Value {
  if (args.length < 1 || args.length > 2) return ERROR_VALUE;
  const info = toText(scalarArg(args[0], state));
  if (typeof info !== 'string') return info;
  const context = state.context;
  let sheet = context.sheet;
  let row = context.row;
  let col = context.col;
  if (args.length > 1 && args[1]) {
    const ref = refArg(args[1], state);
    if (!isRef(ref)) return ref;
    sheet = ref.sheet;
    row = ref.range.top;
    col = ref.range.left;
  }
  switch (info.toLowerCase()) {
    case 'row':
      return row + 1;
    case 'col':
      return col + 1;
    case 'address': {
      const address = `$${columnName(col)}$${row + 1}`;
      const name = sheet === context.sheet ? undefined : context.sheetName(sheet);
      return name === undefined ? address : `${quoteSheetName(name)}!${address}`;
    }
    case 'contents': {
      const value = context.cellValue(sheet, row, col);
      return value === null ? 0 : value;
    }
  }
  return ERROR_VALUE;
}

// ---------------------------------------------------------------------------
// LET, LAMBDA and the functions that take a LAMBDA
//
// A LAMBDA is a value only inside the formula that makes it: it can be given a name with LET,
// be called, or be given to MAP and its relatives. Everywhere else (a cell, an operator, a
// function of the library) it is #CALC!. The names of LET and the parameters of a LAMBDA are
// seen by everything inside them, before the defined names of the workbook.

/** A LAMBDA, or any other value. */
type Callable = Value | LambdaValue;

/** A LAMBDA where only a value can be: #CALC!. */
const valueOf = (value: Callable): Value => (isLambda(value) ? ERROR_CALC : value);

/** Whether the arguments of LAMBDA are distinct names followed by a body. */
function lambdaShape(args: Args): boolean {
  if (args.length < 1 || args.length - 1 > MAX_PARAMS || !args[args.length - 1]) return false;
  const seen: string[] = [];
  for (let k = 0; k < args.length - 1; k++) {
    const param = args[k];
    if (!param || param.type !== 'name' || param.sheet !== undefined) return false;
    const upper = param.name.toUpperCase();
    if (seen.indexOf(upper) >= 0) return false;
    seen.push(upper);
  }
  return true;
}

function makeLambda(args: Args, state: State): LambdaValue | CalcError {
  if (!lambdaShape(args)) return ERROR_VALUE;
  const params: string[] = [];
  for (let k = 0; k < args.length - 1; k++) params.push((args[k] as Extract<Node, { type: 'name' }>).name.toUpperCase());
  return { lambda: true, params, body: args[args.length - 1] as Node, scope: state.scope, rowShift: state.rowShift, colShift: state.colShift };
}

/**
 * A tree where a LAMBDA may come out: LAMBDA itself, a name bound to one, a defined name that
 * stands for one, a call of a function that gives one. Anything else is evaluated as usual.
 */
function evalCallable(node: Node, state: State): Callable {
  if (node.type === 'name' && node.sheet === undefined) {
    const local = scoped(node.name, state);
    if (local) return isOmitted(local.value) ? null : local.value;
    const target = definedTree(node.name, state);
    if (target === undefined) return evalNode(node, state);
    // A name that stands for a name that stands for a name ... is as deep as calls inside one another.
    if (depth - state.base >= MAX_NODE_DEPTH) return ERROR_VALUE;
    depth++;
    const result = withName(node.name, state, () => evalCallable(target, state));
    depth--;
    return result;
  }
  if (node.type !== 'call') return evalNode(node, state);
  if (depth - state.base >= MAX_NODE_DEPTH) return ERROR_VALUE;
  depth++;
  let result: Callable;
  if (node.name === 'LAMBDA') result = makeLambda(node.args, state);
  else if (node.name === 'LET') result = fnLet(node.args, state, true);
  else if (node.name === APPLY_CALL) result = fnApply(node.args, state);
  else if (SPECIALS.has(node.name) || state.functions.has(node.name)) result = evalCall(node, state);
  else {
    const made = calledLambda(node.name, state);
    result = made === undefined ? ERROR_NAME : isLambda(made) ? invoke(made, node.args, state) : made;
  }
  depth--;
  return result;
}

/** The tree of a defined name as seen from the sheet of the formula, or undefined. */
function definedTree(name: string, state: State): Node | undefined {
  const resolve = state.options.resolveName;
  return resolve ? resolve(name, state.context.sheet) : undefined;
}

/**
 * Runs `run` for the tree of a defined name: relative references of the formula do not move
 * it, it sees none of the names of LET, and a name that leads back to itself has no value.
 */
function withName(name: string, state: State, run: () => Callable): Callable {
  const key = `${state.context.sheet}!${name.toUpperCase()}`;
  if (state.names.indexOf(key) >= 0) return ERROR_NAME;
  const rowShift = state.rowShift;
  const colShift = state.colShift;
  const scope = state.scope;
  state.names.push(key);
  state.rowShift = 0;
  state.colShift = 0;
  state.scope = undefined;
  try {
    return run();
  } finally {
    state.rowShift = rowShift;
    state.colShift = colShift;
    state.scope = scope;
    state.names.pop();
  }
}

/** What a name called as a function stands for: a name of LET or a defined name; undefined when it is neither. */
function calledLambda(name: string, state: State): Callable | undefined {
  const local = scoped(name, state);
  if (local) return isOmitted(local.value) ? ERROR_VALUE : local.value;
  const target = definedTree(name, state);
  if (target === undefined) return undefined;
  return withName(name, state, () => evalCallable(target, state));
}

/** Calls a LAMBDA with the arguments written in a formula; an argument left empty is omitted. */
function invoke(lambda: LambdaValue, nodes: Args, state: State): Callable {
  if (nodes.length > lambda.params.length) return ERROR_VALUE;
  const values: Bound[] = new Array(nodes.length);
  for (let k = 0; k < nodes.length; k++) {
    const node = nodes[k];
    values[k] = node ? evalCallable(node, state) : OMITTED;
  }
  return call(lambda, values, state);
}

/** Calls a LAMBDA with values; a parameter without one is omitted. */
function call(lambda: LambdaValue, values: readonly Bound[], state: State): Callable {
  if (values.length > lambda.params.length) return ERROR_VALUE;
  if (state.lambdas >= MAX_LAMBDA_DEPTH || depth - state.base >= MAX_NODE_DEPTH) return ERROR_NUM;
  if (state.context.tick) state.context.tick(1);
  let scope = lambda.scope;
  for (let k = 0; k < lambda.params.length; k++) scope = { name: lambda.params[k], value: k < values.length ? values[k] : OMITTED, parent: scope };
  const savedScope = state.scope;
  const rowShift = state.rowShift;
  const colShift = state.colShift;
  // Like a defined name, the body of a LAMBDA does not see the names of LET where it is called.
  state.scope = scope;
  state.rowShift = lambda.rowShift;
  state.colShift = lambda.colShift;
  state.lambdas++;
  depth++;
  try {
    return evalCallable(lambda.body, state);
  } finally {
    depth--;
    state.lambdas--;
    state.scope = savedScope;
    state.rowShift = rowShift;
    state.colShift = colShift;
  }
}

/** LET(name1, value1, [name2, value2, ...], calculation): each name is seen by what follows it. */
function fnLet(args: Args, state: State, callable: boolean): Callable {
  if (args.length < 3 || args.length % 2 === 0 || (args.length - 1) / 2 > MAX_PARAMS) return ERROR_VALUE;
  const saved = state.scope;
  try {
    for (let k = 0; k + 1 < args.length; k += 2) {
      const name = args[k];
      const value = args[k + 1];
      if (!name || name.type !== 'name' || name.sheet !== undefined || !value) return ERROR_VALUE;
      state.scope = { name: name.name.toUpperCase(), value: evalCallable(value, state), parent: state.scope };
    }
    const body = args[args.length - 1];
    if (!body) return ERROR_VALUE;
    return callable ? evalCallable(body, state) : evalNode(body, state);
  } finally {
    state.scope = saved;
  }
}

/** LAMBDA(x,x+1)(2): the first argument is what is called. */
function fnApply(args: Args, state: State): Callable {
  const callee = args.length > 0 && args[0] ? evalCallable(args[0], state) : ERROR_VALUE;
  if (!isLambda(callee)) return isError(callee) ? callee : ERROR_VALUE;
  return invoke(callee, args.slice(1), state);
}

function fnIsOmitted(args: Args, state: State): Value {
  if (args.length !== 1) return ERROR_VALUE;
  const arg = args[0];
  if (!arg || arg.type !== 'name' || arg.sheet !== undefined) return false;
  const local = scoped(arg.name, state);
  return !!local && isOmitted(local.value);
}

/** The LAMBDA given as the last argument of MAP and its relatives, which must take `count` parameters. */
function lambdaArg(arg: Node | null | undefined, count: number, state: State): LambdaValue | CalcError {
  if (!arg) return ERROR_VALUE;
  const value = evalCallable(arg, state);
  if (!isLambda(value)) return isError(value) ? value : ERROR_VALUE;
  return value.params.length === count ? value : ERROR_VALUE;
}

/** An argument of MAP and its relatives as an array: the cells of a reference, a scalar as an array of one. */
function arrayArgOf(arg: Node | null | undefined, state: State): ArrayValue | CalcError {
  if (!arg) return ERROR_VALUE;
  const value = derefAll(evalArrayArg(arg, state), state);
  if (isArray(value)) return value;
  return isError(value) ? value : { rows: 1, cols: 1, data: [value] };
}

/** What one call of the LAMBDA of MAP, SCAN, BYROW, BYCOL or MAKEARRAY gives for one element: one value, an array of arrays being #CALC!. */
function elementOf(value: Callable, state: State): Scalar {
  if (isLambda(value)) return ERROR_CALC;
  if (isRef(value)) return isSingleCell(value) ? state.context.cellValue(value.sheet, value.range.top, value.range.left) : ERROR_CALC;
  if (isArray(value)) return value.rows === 1 && value.cols === 1 ? value.data[0] : ERROR_CALC;
  return value;
}

const tooLarge = (rows: number, cols: number) => !(rows * cols <= MAX_ARRAY_CELLS);

/** MAP(array1, [array2, ...], lambda): the LAMBDA of the elements at each position. */
function fnMap(args: Args, state: State): Value {
  if (args.length < 2 || args.length - 1 > MAX_PARAMS) return ERROR_VALUE;
  const arrays: ArrayValue[] = [];
  for (let k = 0; k < args.length - 1; k++) {
    const array = arrayArgOf(args[k], state);
    if (!isArray(array)) return array;
    arrays.push(array);
  }
  const lambda = lambdaArg(args[args.length - 1], arrays.length, state);
  if (!isLambda(lambda)) return lambda;
  let rows = 1;
  let cols = 1;
  for (const array of arrays) {
    rows = Math.max(rows, array.rows);
    cols = Math.max(cols, array.cols);
  }
  if (tooLarge(rows, cols)) return ERROR_NUM;
  const data: Scalar[] = new Array(rows * cols);
  const values: Scalar[] = new Array(arrays.length);
  for (let r = 0, k = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++, k++) {
      for (let a = 0; a < arrays.length; a++) values[a] = arrayAt(arrays[a], r, c);
      const value = elementOf(call(lambda, values, state), state);
      if (typeof value === 'string' && !chargeText(value)) return ERROR_VALUE;
      data[k] = value;
    }
  }
  return { rows, cols, data };
}

/** REDUCE(initial, array, lambda(accumulator, value)), and SCAN, which gives every accumulator on the way. */
function fnReduce(args: Args, state: State, scan: boolean): Value {
  if (args.length !== 3) return ERROR_VALUE;
  let accumulator: Value = args[0] ? derefAll(evalNode(args[0], state), state) : null;
  const array = arrayArgOf(args[1], state);
  if (!isArray(array)) return array;
  const lambda = lambdaArg(args[2], 2, state);
  if (!isLambda(lambda)) return lambda;
  const count = array.rows * array.cols;
  const data: Scalar[] = scan ? new Array(count) : [];
  for (let k = 0; k < count; k++) {
    const next = call(lambda, [accumulator, array.data[k]], state);
    if (isLambda(next)) return ERROR_CALC;
    accumulator = derefAll(next, state);
    if (scan) {
      const value = elementOf(accumulator, state);
      if (typeof value === 'string' && !chargeText(value)) return ERROR_VALUE;
      data[k] = value;
    }
  }
  return scan ? { rows: array.rows, cols: array.cols, data } : accumulator;
}

/** BYROW(array, lambda(row)) and BYCOL: one value for each row (or column). */
function fnByLine(args: Args, state: State, byRow: boolean): Value {
  if (args.length !== 2) return ERROR_VALUE;
  const array = arrayArgOf(args[0], state);
  if (!isArray(array)) return array;
  const lambda = lambdaArg(args[1], 1, state);
  if (!isLambda(lambda)) return lambda;
  const lines = byRow ? array.rows : array.cols;
  const length = byRow ? array.cols : array.rows;
  const data: Scalar[] = new Array(lines);
  for (let i = 0; i < lines; i++) {
    const line: Scalar[] = new Array(length);
    for (let j = 0; j < length; j++) line[j] = byRow ? array.data[i * array.cols + j] : array.data[j * array.cols + i];
    if (state.context.tick) state.context.tick(length);
    const value = elementOf(call(lambda, [byRow ? { rows: 1, cols: length, data: line } : { rows: length, cols: 1, data: line }], state), state);
    if (typeof value === 'string' && !chargeText(value)) return ERROR_VALUE;
    data[i] = value;
  }
  return byRow ? { rows: lines, cols: 1, data } : { rows: 1, cols: lines, data };
}

/** MAKEARRAY(rows, columns, lambda(row, column)). */
function fnMakeArray(args: Args, state: State): Value {
  if (args.length !== 3) return ERROR_VALUE;
  const rows = intArg(args[0], state);
  if (typeof rows !== 'number') return rows;
  const cols = intArg(args[1], state);
  if (typeof cols !== 'number') return cols;
  if (rows < 1 || cols < 1) return ERROR_VALUE;
  // UNSURE: what Excel gives for an array too large to make; #NUM!, as SEQUENCE gives here.
  if (tooLarge(rows, cols)) return ERROR_NUM;
  const lambda = lambdaArg(args[2], 2, state);
  if (!isLambda(lambda)) return lambda;
  const data: Scalar[] = new Array(rows * cols);
  for (let r = 0, k = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++, k++) {
      const value = elementOf(call(lambda, [r + 1, c + 1], state), state);
      if (typeof value === 'string' && !chargeText(value)) return ERROR_VALUE;
      data[k] = value;
    }
  }
  return { rows, cols, data };
}

/** ANCHORARRAY(A1), written "A1#": the cells the dynamic array formula of A1 spills into. */
function fnAnchorArray(args: Args, state: State): Value {
  if (args.length !== 1 || !args[0]) return ERROR_VALUE;
  const ref = refArg(args[0], state);
  if (!isRef(ref)) return ref;
  if (!isSingleCell(ref)) return ERROR_REF;
  const ask = state.context.spillRange;
  const range = ask ? ask.call(state.context, ref.sheet, ref.range.top, ref.range.left) : undefined;
  if (range === undefined) return ERROR_REF;
  if (isError(range)) return range;
  return { ref: true, sheet: ref.sheet, range };
}

const SPECIALS = new Map<string, Special>([
  ['IF', fnIf],
  ['IFERROR', (args, state) => ifCaught(args, state, isError)],
  ['IFNA', (args, state) => ifCaught(args, state, (value) => isError(value) && value.error === '#N/A')],
  ['IFS', fnIfs],
  ['SWITCH', fnSwitch],
  ['CHOOSE', fnChoose],
  ['AND', (args, state) => logical(args, state, (trues, count) => trues === count)],
  ['OR', (args, state) => logical(args, state, (trues) => trues > 0)],
  ['XOR', (args, state) => logical(args, state, (trues) => trues % 2 === 1)],
  ['NOT', fnNot],
  ['ROW', (args, state) => rowOrColumn(args, state, true)],
  ['COLUMN', (args, state) => rowOrColumn(args, state, false)],
  ['ROWS', (args, state) => rowsOrColumns(args, state, true)],
  ['COLUMNS', (args, state) => rowsOrColumns(args, state, false)],
  ['AREAS', fnAreas],
  ['OFFSET', fnOffset],
  ['INDIRECT', fnIndirect],
  ['INDEX', fnIndex],
  ['ISREF', (args, state) => (args.length !== 1 ? ERROR_VALUE : isRef(evalArg(args[0], state)))],
  ['ISFORMULA', fnIsFormula],
  ['FORMULATEXT', fnFormulaText],
  ['SHEET', fnSheet],
  ['SHEETS', fnSheets],
  ['ISBLANK', fnIsBlank],
  ['N', fnN],
  ['T', fnT],
  ['CELL', fnCell],
  // The "@" of newer versions is stored as a function: implicit intersection, asked for by name.
  ['SINGLE', (args, state) => (args.length !== 1 ? ERROR_VALUE : scalarArg(args[0], state))],
  // "A1#" is stored as a function: the cells the dynamic array formula of A1 spills into.
  ['ANCHORARRAY', fnAnchorArray],
  ['LET', (args, state) => valueOf(fnLet(args, state, false))],
  // A function is no value a cell can hold.
  ['LAMBDA', (args) => (lambdaShape(args) ? ERROR_CALC : ERROR_VALUE)],
  [APPLY_CALL, (args, state) => valueOf(fnApply(args, state))],
  ['ISOMITTED', fnIsOmitted],
  ['MAP', fnMap],
  ['REDUCE', (args, state) => fnReduce(args, state, false)],
  ['SCAN', (args, state) => fnReduce(args, state, true)],
  ['BYROW', (args, state) => fnByLine(args, state, true)],
  ['BYCOL', (args, state) => fnByLine(args, state, false)],
  ['MAKEARRAY', fnMakeArray],
]);

/** Names the evaluator answers itself; the function library is not asked for them. */
export function isBuiltinForm(name: string): boolean {
  return SPECIALS.has(name);
}
