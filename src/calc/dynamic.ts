// Whether a formula is a dynamic array formula: one that means something else written as an
// ordinary formula (with implicit intersection), so that it is written to a file as an array
// formula with the cell metadata of a dynamic one.
//
// The answer comes from the text alone, never from what the formula is worth at a moment: a
// formula typed in and one read from the file it was saved to are the same formula, for
// everyone who edits the workbook together. Where the text does not say for certain, the
// formula is taken for a dynamic one.
//
// UNSURE: Excel's own rule was not checked against files it wrote. Its documentation says a
// formula is written as a dynamic array formula when it could return several values or would
// use implicit intersection written as an ordinary one; this is that rule worked out from the
// tree, and errs on the side of a dynamic formula (which means the same either way).

import { FUNCTIONS } from './functions/index';
import { APPLY_CALL, parseFormula, type Node } from './parse';
import type { ArgSelector } from './values';

/** Functions whose result may be several values, wherever they stand but inside an argument that takes an array. */
const ARRAY_RESULT = new Set([
  'FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'SEQUENCE', 'RANDARRAY', 'TRANSPOSE', 'MMULT', 'MUNIT', 'FREQUENCY',
  'LINEST', 'LOGEST', 'TREND', 'GROWTH', 'TEXTSPLIT', 'VSTACK', 'HSTACK', 'TAKE', 'DROP', 'CHOOSECOLS', 'CHOOSEROWS',
  'TOCOL', 'TOROW', 'WRAPROWS', 'WRAPCOLS', 'EXPAND',
]);

/** Functions that give a reference, of several cells perhaps: fine where a range is taken as it is. */
const REF_RESULT = new Set(['OFFSET', 'INDIRECT', 'XLOOKUP', 'ROW', 'COLUMN']);

/** Forms that exist only with dynamic arrays, wherever they stand: names of their own, LAMBDA and what takes one, "#". */
const DYNAMIC_ONLY = new Set(['LET', 'LAMBDA', 'MAP', 'REDUCE', 'SCAN', 'BYROW', 'BYCOL', 'MAKEARRAY', 'ISOMITTED', 'ANCHORARRAY', APPLY_CALL]);

/** Forms of the evaluator whose arguments are taken as ranges (asked where they are, or read cell by cell). */
const RANGE_ARGS = new Set(['ROWS', 'COLUMNS', 'AREAS', 'ISREF', 'SHEET', 'SHEETS', 'ISFORMULA', 'FORMULATEXT', 'CELL', 'AND', 'OR', 'XOR', 'ROW', 'COLUMN']);

/** Forms of the evaluator that take single values (an ordinary formula intersects a range given to them). */
const SCALAR_FORMS = new Set(['IF', 'IFERROR', 'IFNA', 'IFS', 'SWITCH', 'CHOOSE', 'NOT', 'ISBLANK', 'N', 'T', 'OFFSET', 'INDIRECT']);

const selects = (selector: ArgSelector | undefined, index: number): boolean => selector === true || (selector !== undefined && selector(index) === true);

const MAX_NODES = 100000;

// Where a node stands, as an ordinary formula reads what is there.
/** One value is taken: a range is intersected, an array gives its first value. */
const SCALAR = 0;
/** A reference is taken as it is (the range of SUM, of SUMIF); anything an operator works on is one value. */
const RANGED = 1;
/** An argument that takes an array (SUMPRODUCT, the first of INDEX): an ordinary formula reads it as an array formula does. */
const ARRAY = 2;

const isConstantIndex = (node: Node | null | undefined): boolean => !!node && node.type === 'number' && node.value >= 1;

const isOneCell = (node: Node): boolean => node.type === 'ref' && !node.wholeRows && !node.wholeCols && node.range.top === node.range.bottom && node.range.left === node.range.right;

/** Whether the formula, without "=", is a dynamic array formula by the rule above. A text that cannot be parsed is none. */
export function needsDynamic(text: string): boolean {
  let root: Node;
  try {
    root = parseFormula(text);
  } catch (error) {
    return false;
  }
  return nodeNeedsDynamic(root);
}

export function nodeNeedsDynamic(root: Node): boolean {
  const nodes: Node[] = [root];
  const places: number[] = [SCALAR];
  for (let seen = 0; nodes.length > 0; seen++) {
    if (seen > MAX_NODES) return true;
    const node = nodes.pop() as Node;
    const place = places.pop() as number;
    const push = (child: Node | null | undefined, at: number) => {
      if (child) {
        nodes.push(child);
        places.push(at);
      }
    };
    switch (node.type) {
      case 'array':
        if (place !== ARRAY && node.rows.length * (node.rows[0] ? node.rows[0].length : 0) > 1) return true;
        break;
      case 'ref':
        if (place === SCALAR && !isOneCell(node)) return true;
        break;
      case 'name':
        // A defined name may be a range, and what it stands for can change.
        if (place === SCALAR) return true;
        break;
      case 'unary':
        push(node.operand, place === ARRAY ? ARRAY : SCALAR);
        break;
      case 'binary': {
        const joins = node.op === ':' || node.op === ' ' || node.op === ',';
        // A range put together, an intersection or a union may be several cells; an operator takes one value of each side.
        if (joins && place === SCALAR) return true;
        const at = place === ARRAY ? ARRAY : joins ? RANGED : SCALAR;
        push(node.left, at);
        push(node.right, at);
        break;
      }
      case 'call': {
        const name = node.name;
        if (DYNAMIC_ONLY.has(name)) return true;
        if (name === 'SINGLE') {
          // "@": one value, whatever stands in it.
          for (const arg of node.args) push(arg, ARRAY);
          break;
        }
        if (place === ARRAY) {
          // As an array formula reads it: only what exists with dynamic arrays alone matters.
          for (const arg of node.args) push(arg, ARRAY);
          break;
        }
        if (ARRAY_RESULT.has(name)) return true;
        if (place === SCALAR && REF_RESULT.has(name)) {
          // ROW(A1) and COLUMN(A1) of one cell give one number.
          if (!((name === 'ROW' || name === 'COLUMN') && (node.args.length === 0 || (node.args[0] && isOneCell(node.args[0]))))) return true;
        }
        if (name === 'INDEX') {
          // One cell when both its row and column are given; a row or a column of it otherwise, perhaps.
          if (place === SCALAR && !(isConstantIndex(node.args[1]) && isConstantIndex(node.args[2]))) return true;
          push(node.args[0], ARRAY);
          for (let k = 1; k < node.args.length; k++) push(node.args[k], SCALAR);
          break;
        }
        const spec = FUNCTIONS.get(name);
        if (RANGE_ARGS.has(name)) {
          for (const arg of node.args) push(arg, RANGED);
          break;
        }
        if (SCALAR_FORMS.has(name)) {
          for (const arg of node.args) push(arg, SCALAR);
          break;
        }
        // A function this engine does not know, or a defined name that stands for a LAMBDA.
        if (!spec) return true;
        for (let k = 0; k < node.args.length; k++) push(node.args[k], selects(spec.arrayArgs, k) ? ARRAY : selects(spec.scalarArgs, k) ? SCALAR : RANGED);
        break;
      }
    }
  }
  return false;
}
