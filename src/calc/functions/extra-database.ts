// Database functions: DSUM, DCOUNT, DCOUNTA, DAVERAGE, DMAX, DMIN and DGET.
//
// A database is a range whose first row names the fields; every row below is
// a record. The criteria range has a row of field names too; each row below it
// is one alternative (rows are OR), and the filled cells of a row must all hold
// (columns are AND). An empty criteria cell asks for nothing.

import {
  ERROR_DIV0,
  ERROR_NUM,
  ERROR_VALUE,
  isError,
  parseErrorCode,
  toText,
  type CalcContext,
  type CalcFunction,
  type FunctionSpec,
  type Scalar,
} from '../values';
import {
  clean,
  compileCriterion,
  foldCase,
  mapScalarsExcept,
  parseNumberText,
  spec,
  toArray,
  toGrid,
  usedCols,
  usedRows,
  type Criterion,
} from './helpers';

/** Tests (records x conditions) a call may make before it gives up with #NUM!. */
const MAX_TESTS = 1 << 26;

type Kind = 'sum' | 'count' | 'counta' | 'average' | 'max' | 'min' | 'get';

interface Condition {
  field: number;
  criterion: Criterion;
}

function startsWithOperator(text: string): boolean {
  const c = text.charCodeAt(0);
  // = < >
  return c === 61 || c === 60 || c === 62;
}

/**
 * The criterion of one criteria cell. It reads as the criteria of COUNTIF do,
 * with one difference: text written without an operator finds every text that
 * begins with it ("Dav" finds "David"); `="=Dav"` asks for exactly "Dav".
 */
function cellCriterion(cell: Scalar, context: CalcContext): Criterion {
  if (typeof cell === 'string' && !startsWithOperator(cell)) {
    const upper = cell.toUpperCase();
    const plain =
      parseNumberText(cell, context) === undefined &&
      upper !== 'TRUE' &&
      upper !== 'FALSE' &&
      parseErrorCode(cell) === undefined;
    if (plain) {
      // A "~" left alone at the end would take the added "*" literally.
      let tildes = 0;
      for (let i = cell.length - 1; i >= 0 && cell.charCodeAt(i) === 126; i--) tildes++;
      return compileCriterion(cell + (tildes % 2 === 1 ? '~*' : '*'), context);
    }
  }
  return compileCriterion(cell, context);
}

function headerText(value: Scalar): string {
  const text = toText(value);
  return isError(text) ? '' : foldCase(text);
}

function database(kind: Kind): CalcFunction {
  return mapScalarsExcept(
    (index) => index !== 1,
    (args, context) => {
      const source = args[0];
      if (isError(source)) return source;
      if (isError(args[2])) return args[2];
      const table = toGrid(source, context);
      if (table.rows < 1 || table.cols < 1) return ERROR_VALUE;
      const rows = usedRows(source, table, context);
      const cols = usedCols(source, table, context);
      const headers: string[] = new Array(cols);
      for (let c = 0; c < cols; c++) headers[c] = headerText(table.get(0, c));
      const fieldOf = (name: string): number => {
        if (name === '') return -1;
        for (let c = 0; c < cols; c++) if (headers[c] === name) return c;
        return -1;
      };

      // The field: its name or its position from 1; only the counting functions may leave it out.
      const given = args.length > 1 ? clean(args[1] as Scalar) : null;
      let field = -1;
      if (isError(given)) return given;
      if (given === null) {
        if (kind !== 'count' && kind !== 'counta') return ERROR_VALUE;
      } else if (typeof given === 'string') {
        field = fieldOf(foldCase(given));
        if (field < 0) return ERROR_VALUE;
      } else if (typeof given === 'number') {
        field = Math.trunc(given) - 1;
        if (!(field >= 0 && field < table.cols)) return ERROR_VALUE;
      } else return ERROR_VALUE;

      const criteria = toArray(args[2], context);
      if (isError(criteria)) return criteria;
      if (criteria.rows < 2) return ERROR_VALUE;
      const fields: number[] = new Array(criteria.cols);
      for (let c = 0; c < criteria.cols; c++) fields[c] = fieldOf(headerText(clean(criteria.data[c])));
      const alternatives: Condition[][] = [];
      let everything = false;
      let tests = 0;
      for (let r = 1; r < criteria.rows && !everything; r++) {
        const conditions: Condition[] = [];
        for (let c = 0; c < criteria.cols; c++) {
          const cell = clean(criteria.data[r * criteria.cols + c]);
          if (cell === null || cell === '') continue;
          // A condition under a heading that names no field is a calculated criterion: it would have
          // to be worked out again for every record, which the value in the cell does not allow.
          if (fields[c] < 0) return ERROR_VALUE;
          conditions.push({ field: fields[c], criterion: cellCriterion(cell, context) });
        }
        if (conditions.length === 0) everything = true;
        else {
          alternatives.push(conditions);
          tests += conditions.length;
        }
      }
      if (!everything && tests * rows > MAX_TESTS) return ERROR_NUM;

      let total = 0;
      let count = 0;
      let extreme = 0;
      let matched = 0;
      let found: Scalar = null;
      for (let r = 1; r < rows; r++) {
        let match = everything;
        for (let a = 0; a < alternatives.length && !match; a++) {
          const conditions = alternatives[a];
          match = true;
          for (let i = 0; i < conditions.length && match; i++) {
            match = conditions[i].criterion.test(table.get(r, conditions[i].field));
          }
        }
        if (!match) continue;
        matched++;
        if (field < 0) continue;
        const value = table.get(r, field);
        if (kind === 'get') {
          if (matched > 1) return ERROR_NUM;
          found = value;
        } else if (kind === 'counta') {
          if (value !== null) count++;
        } else if (typeof value === 'number') {
          total += value;
          if (count === 0 || (kind === 'max' ? value > extreme : value < extreme)) extreme = value;
          count++;
        } else if (isError(value) && kind !== 'count') return value;
      }
      // A database cut to the used part of its sheet may have empty records below; with criteria
      // that ask for nothing they match as well, but they hold no value to add.
      if (everything && rows < table.rows) matched += table.rows - Math.max(rows, 1);
      switch (kind) {
        case 'sum':
          return isFinite(total) ? total : ERROR_NUM;
        case 'count':
        case 'counta':
          return field < 0 ? matched : count;
        case 'average':
          return count === 0 ? ERROR_DIV0 : isFinite(total / count) ? total / count : ERROR_NUM;
        case 'get':
          return matched === 0 ? ERROR_VALUE : matched > 1 ? ERROR_NUM : found;
        default:
          return count === 0 ? 0 : extreme;
      }
    },
  );
}

export const EXTRA_DATABASE_FUNCTIONS: { [name: string]: FunctionSpec } = {
  DSUM: spec(database('sum'), 3, 3),
  DCOUNT: spec(database('count'), 3, 3),
  DCOUNTA: spec(database('counta'), 3, 3),
  DAVERAGE: spec(database('average'), 3, 3),
  DMAX: spec(database('max'), 3, 3),
  DMIN: spec(database('min'), 3, 3),
  DGET: spec(database('get'), 3, 3),
};
