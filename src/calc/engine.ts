// Recalculation of a workbook.
//
// The engine keeps, for every formula, the cells it reads (its precedents). When cells
// change, the formulas that read them are found through that index, then the formulas that
// read those, and only this set is calculated. Formulas outside it keep the value stored in
// the model: what a file says its formulas are worth is trusted, a workbook is not
// recalculated when it is opened.
//
// Ranges are kept as ranges. A formula that reads a whole column is one entry in the index,
// not a million.
//
// Three kinds of formula are left alone, with the value the file stores for them: those
// whose text cannot be parsed, those that use something this engine does not have (an
// unknown function, a reference across sheets or into another workbook, a use of a function
// that is only partly supported), and those that cannot be calculated within the budget of
// a recalculation. A fourth kind is calculated but keeps its value: a formula that is
// calculated again although nothing it reads changed, and comes out different from what the
// file says. The file was written by a program that knew something this engine does not.
//
// Everything that can take long is done in slices: reading the formulas into the graph,
// finding the circular references, calculating. A call that runs out of budget returns with
// `incomplete` set and the next call goes on where it stopped.
//
// Two things a formula can depend on are not cells: whether rows are hidden (SUBTOTAL,
// AGGREGATE) and what the defined names stand for. The formulas that use them are noted
// when the graph is built, and rowsChanged() and namesChanged() name them as changed.

import { touchModel } from '../model/edits';
import { isoToSerial } from '../format/cell-format';
import type { FormatLocale } from '../format/locale';
import { MAX_COLS, MAX_ROWS, parseRange, rangeName, type Cell, type CellFormula, type CellRange, type Row, type Sheet, type Workbook } from '../model/sheet/types';
import { sharedMasters, type SharedMasters } from '../ops/ops';
import { EvaluationAbort, evaluate, evaluationDepth, isBuiltinForm, refRange, toCellValue, unionParts, type EvaluateOptions } from './eval';
import { FUNCTIONS } from './functions/index';
import { parseFormula, type Node } from './parse';
import { shiftFormula } from './refs';
import { ERROR_NUM, ERROR_VALUE, arrayAt, calcError, isArray, isError, isRef, isWholeLine, parseErrorCode, readRange, refToArray, type ArrayValue, type CalcContext, type CalcError, type FunctionSpec, type Scalar, type Value } from './values';

/** How much one recalculate() call may cost before it stops with part of the formulas left for the next call. */
export interface CalcBudget {
  /** Milliseconds on the clock. 5000 when omitted; 0 or Infinity for no limit. */
  milliseconds?: number;
  /**
   * Units of work: cells read, elements of arrays calculated, functions called. No limit when
   * omitted. Only calculating counts; building the dependency graph is limited by `milliseconds`.
   */
  work?: number;
  /**
   * What a formula may take that cannot be calculated within `milliseconds` even with a call
   * to itself: the call goes on until it is done or this limit is reached, and only then is
   * the formula given up (see CalcEngine.overBudget). Ten times `milliseconds` and at least
   * 2000 when omitted; 0 or Infinity for no limit.
   */
  hardMilliseconds?: number;
  /** The same for `work`: ten times `work` and at least 1,000,000 when omitted. */
  hardWork?: number;
}

export interface CalcEngineOptions {
  /** The function library; the built-in one when omitted. */
  functions?: Map<string, FunctionSpec>;
  /** The current date and time as a serial number; read once per recalculation. */
  now?: () => number;
  /** A random number in [0, 1). */
  random?: () => number;
  /**
   * The locale functions follow: the case of letters (UPPER, LOWER, PROPER), and the decimal
   * sign text is read with when it is no number the way files write one (criteria, VALUE,
   * NUMBERVALUE, text met by an operator). Asked every time a
   * formula is calculated, so a view can answer with the locale it shows at that moment.
   * English when omitted. A formula is not calculated again because the answer changed.
   */
  locale?: () => FormatLocale | undefined;
  /** Reserved for iterative calculation, which is not supported: circular references always give 0. */
  maxIterations?: number;
  /**
   * What one recalculate() call may cost. When it is used up the call returns with
   * `incomplete` set: the formulas not reached keep their values and are calculated by
   * the next call.
   */
  budget?: CalcBudget;
  /** Milliseconds from any fixed point, for the budget; Date.now when omitted. */
  clock?: () => number;
  /**
   * Smaller limits for spills than the built-in ones (MAX_SPILL_CELLS, MAX_WORKBOOK_SPILL_CELLS),
   * for tests; a larger value is cut to the built-in one.
   */
  limits?: { spillCells?: number; workbookSpillCells?: number };
}

export interface RecalculateOptions {
  /**
   * The cells are named by the list of sheets the engine read last, which the workbook no
   * longer has: the engine does not look at the list of now and works with the graph it has.
   * For the caller that tells the engine what a sheet that went took along, and calls
   * invalidate() right after (see recalculateForSheets() of the view).
   */
  sheetsOfBefore?: boolean;
  /**
   * With a recalculation of everything: every formula counts as changed, so what is
   * calculated replaces the value of the file also where the two differ. For use after an
   * edit that can change any value without naming cells (rows or columns deleted, a defined
   * name changed). Without it, a formula none of whose precedents changed keeps a value of
   * the file that differs from what the engine calculates (see CalcEngine.mismatched).
   */
  overwrite?: boolean;
}

export interface CellAddress {
  sheet: number;
  row: number;
  col: number;
}

/** What evaluateRange() gives: every value of a reference or an array. */
export interface RangeValues {
  /** Size of the whole result; a reference is cut to the used part of its sheet. */
  rows: number;
  cols: number;
  /** Row by row, null for an empty cell; rows * cols values unless `truncated`. */
  values: Scalar[];
  /** The result has more cells than were asked for: `values` holds the first ones only. */
  truncated: boolean;
  /** Where the values are, when the formula gave a reference: the sheet and the cells (as cut). */
  sheet?: number;
  range?: CellRange;
}

export interface EvaluateRangeOptions {
  /** Most values to give; 1,000,000 when omitted, 5,000,000 at most. */
  maxCells?: number;
  /**
   * The cell the formula was written for, when that is not the cell it is calculated for:
   * relative references move by the distance between the two (see evaluateShifted).
   */
  anchor?: { row: number; col: number };
}

// State of a formula within one recalculation.
const FRESH = 0;
const ACTIVE = 1;
const DONE = 2;

// What becomes of the result of a formula (FormulaRec.keep).
const KEEP_NONE = 0;
/** The same as the value of the file: the cell is left as it is. */
const KEEP_SAME = 1;
/** Different from the value of the file, which stays. */
const KEEP_STORED = 2;

/** Everything the engine knows about one formula cell. */
interface FormulaRec {
  sheet: number;
  row: number;
  col: number;
  /** The formula object of the cell this record was built from; a different object means the cell was changed. */
  formula: CellFormula;
  /** Formula text the tree was parsed from: the master's text for a cell of a shared formula. */
  text: string;
  node: Node;
  /** Distance from the master of a shared formula; relative references move by this much. */
  dRow: number;
  dCol: number;
  /** Calculated whenever anything changed: reads the clock, a random number or cells chosen while calculating. */
  volatile: boolean;
  /** Reads the clock or a random number: its value changes by itself. */
  timed: boolean;
  isArray: boolean;
  /** Ranges are arrays everywhere in the formula: an array formula or a dynamic array formula. */
  arrayMode: boolean;
  /** One of its references names its own cell. */
  selfRef: boolean;
  /** Calls SUBTOTAL or AGGREGATE, itself or through a name: what it is worth depends on which rows are hidden. */
  rowAware: boolean;
  /** The defined names it uses, itself or through other names, in upper case; also those the workbook lacks. */
  names?: string[];
  /** Could not be calculated within the hard budget; left alone until something it reads changes. */
  costly: boolean;
  /** Cells an array formula fills, when there is more than one. */
  spill?: CellRange;
  /**
   * A dynamic array formula: calculated as an array formula, its result spills from its cell
   * into as many cells as it has, or it is #SPILL! when something is in the way.
   */
  dynamic: boolean;
  /** Dynamic: what the last calculation found (see run()), for conclude() to apply. */
  next?: SpillPlan;
  /** Dynamic: the cells the result is meant to fill, which the formula watches for what may be in the way (a precedent of `precedents`). */
  guard?: Precedent;
  /** Times the formula was taken up again in one recalculation because the cells of a spill it read changed (see reopen()). */
  reopens: number;
  reopenEpoch: number;
  precedents: Precedent[];
  /** The recalculation this record is part of; anything else means "not dirty". */
  epoch: number;
  state: number;
  value: Scalar;
  /** Whole result of an array formula. */
  result?: ArrayValue;
  /** Part of a circular reference made of the references the formulas name. */
  circular: boolean;
  /** Part of a circular reference found while calculating (through OFFSET, INDIRECT). */
  looped: boolean;
  /** Something it reads, or its own formula, changed: its result replaces whatever the cell holds. */
  live: boolean;
  /** KEEP_*: what the last calculation decided about the value the cell holds. */
  keep: number;
  /** Listed in `mismatched`. */
  kept: boolean;
  /** Dynamic: the cells spills it gave up since its result was last written covered (see store()). */
  left?: CellRange;
  /** Found circular when its spills were settled (see resolve()): 0 in its own cell until it is calculated anew. */
  forced?: boolean;
  forcedEpoch?: number;
  /** The formulas found circular together with it (see resolve()). */
  forcedGroup?: FormulaRec[];
  /** How far the whole columns and rows it reads went when it was last looked at (see begin()). */
  /** A spill it read changed while it was being calculated: it is calculated again (see takeUp()). */
  stale?: boolean;
  /** The cell or its formula changed without the engine being told, or the record was taken out of the graph. */
  gone: boolean;
  /** Index in the chain of formulas being calculated. */
  position: number;
  // Finding circular references (see components()).
  stamp: number;
  index: number;
  low: number;
  onStack: boolean;
  /** Group of formulas that read one another, 0 for none. */
  comp: number;
  /** Where the record is in the column index of its sheet. */
  column?: Column;
  at: number;
}

/** What the calculation of a dynamic array formula found: the cells its result fills, or why it cannot. */
interface SpillPlan {
  /** The cells the result is meant to fill, its own cell first; cut to the sheet. */
  want: CellRange;
  /** The cells it does fill when it spills (more than one); undefined for a single value and when it is blocked. */
  spill?: CellRange;
  /** Something is in the way, or the result does not fit: the formula is #SPILL!. */
  blocked: boolean;
  result?: ArrayValue;
  /** `want` counts in the budget of all spills: it fits and nothing is in its way (see allocate()). */
  budget?: boolean;
}

/** A range that formulas read, as a node of the graph in which circular references are looked for. */
interface RangeNode {
  area: true;
  sheet: number;
  range: CellRange;
  stamp: number;
  index: number;
  low: number;
  onStack: boolean;
}

type GraphNode = FormulaRec | RangeNode;

interface Precedent {
  sheet: number;
  range: CellRange;
  /** Only asked where it is (ROW, COLUMNS, the base of OFFSET): its cells are not read. */
  shapeOnly: boolean;
  /** Set for a range of several cells that is read. */
  node?: RangeNode;
}

interface Entry {
  top: number;
  left: number;
  bottom: number;
  right: number;
  rec: FormulaRec;
}

interface Bucket {
  entries: Entry[];
  /** Recalculation in which every formula of the bucket was already found dirty. */
  done: number;
  /** Recalculation in which every formula of the bucket was already found dirty and live. */
  doneLive: number;
}

/** The formulas of one column in row order, for finding those inside a range. */
interface Column {
  rows: number[];
  recs: FormulaRec[];
  /** For every position, the first position from there on whose formula still matters (see find()). */
  skip: Int32Array;
}

interface SheetState {
  index: number;
  model: Sheet;
  /** The name the sheet had when it was read. */
  name: string;
  recs: Map<number, FormulaRec>;
  /** Formulas that read one cell, by the key of that cell. */
  cellDeps: Map<number, FormulaRec[]>;
  /** Ranges that are not tall, by the bands of rows they touch. */
  rowBands: Map<number, Bucket>;
  /** Tall but narrow ranges, by the columns they touch. */
  colBands: Map<number, Bucket>;
  /** Ranges both tall and wide. */
  wide: Bucket;
  /** Cells filled by array formulas, for ranges small enough to list. */
  spillCells: Map<number, FormulaRec>;
  spillRanges: FormulaRec[];
  /**
   * Listed spills registered over a cell another one had (while spills are settled together, see resolve()): the
   * ones under, by cell, so that the cell is theirs again when the one over it goes.
   */
  shadowed: Map<number, FormulaRec[]>;
  /** Every formula that fills more than its own cell. */
  spills: Set<FormulaRec>;
  /** The ranges of several cells that formulas read on this sheet, one node per range. */
  ranges: Map<string, RangeNode>;
  /** Formulas by column, built when circular references are looked for; undefined after the formulas changed. */
  columns?: Map<number, Column>;
  /** Keys of `columns` in ascending order. */
  columnKeys: number[];
  /**
   * Dynamic array formulas found circular when spills were settled (see resolve()): they spill nowhere, but the circle
   * still runs through the cells they would spill into (their guards), so that it is known for as long as it holds.
   */
  forcedRecs: Set<FormulaRec>;
  /**
   * The cells array formulas outside the graph fill as the file says (`ref`): formulas that cannot
   * be calculated here and keep their values. Nothing spills into them, and what they hold is theirs.
   */
  foreign: Map<number, CellRange>;
  /**
   * How far each column holds something (contentRows()): one past its last row with a value or a formula, by column,
   * built in one pass when first asked for and kept up with the cells the engine is told were changed (noteContent());
   * a column whose last cell was emptied is looked at again (`unsure`). What the engine writes itself (results, the
   * cells of spills) is never content, so it changes none of it.
   */
  lastRows?: Map<number, number>;
  unsure: Set<number>;
  /** One more after every cell the engine was told was changed: what contentCols() found holds until then. */
  contentVersion: number;
  extents?: { version: number; found: Map<string, number> };
  /** The whole columns and rows formulas read on this sheet, with their readers and how far they went (see begin()). */
  wideRanges: Map<string, { range: CellRange; readers: Set<FormulaRec>; seen?: string }>;
  /** The spills on this sheet by the columns they cover (narrow ones), and the wide ones: for reachOf(). */
  spillCols: Map<number, Set<FormulaRec>>;
  wideSpills: Set<FormulaRec>;

  /** Row indexes in ascending order, built when a large range is first read; built again when the rows changed. */
  rowKeys?: number[];
  rowKeysOf?: Map<number, Row>;
  /** Per row, the last recalculation in which a formula of the row was dirty: a cheap first test before a record is looked up. */
  rowEpoch: Uint32Array;
  /** Rows already replaced by a copy in this recalculation. */
  copied: Set<number>;
  changed: { row: number; col: number }[];
}

interface NameEntry {
  text: string;
  /** undefined: not parsed yet; null: cannot be parsed. */
  node?: Node | null;
}

/** Where reading the formulas into the graph stands. */
interface BuildState {
  sheet: number;
  rows?: Iterator<[number, Row]>;
  masters: MastersHolder;
}

/** The masters of the shared formulas of a sheet, worked out when the first one is needed. */
interface MastersHolder {
  masters?: SharedMasters;
}

/** A search for circular references in progress (Tarjan's algorithm with a stack of its own). */
interface ComponentRun {
  /** Every formula is looked at. */
  all: boolean;
  seeds: FormulaRec[];
  next: number;
  stamp: number;
  counter: number;
  /** Nodes whose group is not complete yet. */
  stack: GraphNode[];
  /** The path being walked, with where each node is in the list of what it reads. */
  path: GraphNode[];
  a: number[];
  b: number[];
  c: number[];
  /** For a range: the formulas elsewhere that fill its cells (see successor()). */
  d: (FormulaRec[] | undefined)[];
  steps: number;
  limit: number;
}

/** A changed cell whose formula was read into the graph and whose dependents are still to be marked. */
interface ReadyCell {
  sheet: number;
  row: number;
  col: number;
  /** The cell holds a formula it did not hold before. */
  replaced: boolean;
}

type MutableContext = { -readonly [K in keyof CalcContext]: CalcContext[K] };

const BAND_SHIFT = 7;
/** A range touching more row bands than this is indexed by column, or not at all. */
const MAX_BANDS = 8;
/** Largest range of an array formula that is filled; a larger one (a damaged file) only gets its own cell. */
const MAX_SPILL_CELLS = 1000000;
/**
 * Cells the dynamic array formulas of a workbook fill in all: each of them may spill into an
 * equal share of it (see spillShare()), and is #SPILL! past that, unless the file has it fill those
 * very cells. A spilled cell costs a few hundred bytes. UNSURE: Excel has no such limit (only the
 * size of the sheet); this one keeps a small file from taking gigabytes.
 */
export const MAX_WORKBOOK_SPILL_CELLS = 2000000;
const MAX_LISTED_SPILL = 10000;
/** Formulas waiting on one another on the stack before the engine starts over from the deepest one. */
const MAX_LIVE = 48;
const MAX_LIVE_DEPTH = 300;
/** Below this many rows a range is read row by row instead of through the sorted row list. */
const DIRECT_ROWS = 32;

/** Calculated whenever anything changed. */
const VOLATILE = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'RANDARRAY', 'OFFSET', 'INDIRECT']);

/** They leave hidden rows out: their value changes when rows are hidden or shown. */
const ROW_AWARE = new Set(['SUBTOTAL', 'AGGREGATE']);

/** Their value changes although no cell did. */
const TIMED = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'RANDARRAY']);

/** Functions that only ask where a reference is; the cells of a reference given to them directly are not read. */
const SHAPE_ONLY = new Set(['ROW', 'COLUMN', 'ROWS', 'COLUMNS', 'AREAS', 'SHEET', 'SHEETS', 'ISREF', 'ISFORMULA', 'FORMULATEXT']);

/**
 * Times one formula is taken up again in a round of a recalculation because a spill it reads changed; past it a
 * dynamic array formula is #SPILL! for now, and is tried again with its count started over (see recalculate()).
 */
const MAX_REOPENS = 3;
/** Areas left by spills that are remembered until results are written; past it, what they held is read until then. */
const MAX_VACATED = 256;
/** Cells of array formulas a recalculation makes where the file has none (see store()). */
const MAX_ARRAY_FILL = 100000;
/** Rounds of sharing out the budget of spills again after a formula's lot changed (see allocate()). */
const MAX_ALLOCATE_ROUNDS = 8;
/** Rounds in which the spills of a circle are settled together, from none (see resolve()). */
const MAX_SETTLE_ROUNDS = 16;
/** Times one recalculation looks for the circles its spills made or broke. */
const MAX_CIRCLE_ROUNDS = 8;
/** At most so many dynamic array formulas of a file are put back into the circles they were saved in (see rebuildForced()). */
const MAX_REBUILT = 1000;

/** The functions whose arguments name what they bind: LET binds its names, LAMBDA its parameters. */
const BINDERS = new Set(['LET', 'LAMBDA']);

/** What CELL can tell: the first three only ask where the cell is. */
const CELL_WHERE = new Set(['row', 'col', 'address']);
const CELL_INFO = new Set(['row', 'col', 'address', 'contents']);

/** Values evaluateRange() gives when no limit is asked for, and the most it ever gives. */
const DEFAULT_RANGE_VALUES = 1000000;
const MAX_RANGE_VALUES = 5000000;

const DEFAULT_MILLISECONDS = 5000;
const MIN_HARD_MILLISECONDS = 2000;
const MIN_HARD_WORK = 1000000;
const HARD_FACTOR = 10;
/** Units of work between two looks at the clock. */
const CLOCK_EVERY = 2048;
/** Formulas read into the graph between two looks at the clock. */
const BUILD_EVERY = 256;
/** Steps of the search for circular references between two looks at the clock. */
const SEARCH_EVERY = 4096;
/** Two numbers closer than this, relative to the larger one, are the same value written twice. */
const SAME_NUMBER = 1e-12;

/** Thrown through the evaluator to start over with a shallower stack. */
const UNWIND = new EvaluationAbort();
/** Thrown when the budget of a recalculation is used up. */
const BUDGET = new EvaluationAbort();
const ZERO: ArrayValue = { rows: 1, cols: 1, data: [0] };

const keyOf = (row: number, col: number) => row * MAX_COLS + col;

/** Sheet names compare without regard to case; the Turkish dotted and dotless i count as one letter. */
function foldName(name: string): string {
  return String(name).toLowerCase().replace(/̇/g, '').replace(/ı/g, 'i');
}

function isIndex(value: unknown, limit: number): value is number {
  return typeof value === 'number' && value >= 0 && value < limit && Math.floor(value) === value;
}

/** The value a cell holds in the model. */
function modelValue(cell: Cell, date1904: boolean): Scalar {
  const value = cell.value;
  switch (cell.kind) {
    case 'number':
      return typeof value === 'number' && isFinite(value) ? value : null;
    case 'string':
      return value === null ? null : String(value);
    case 'boolean':
      return typeof value === 'boolean' ? value : null;
    case 'error': {
      const code = parseErrorCode(String(value));
      return code ? calcError(code) : ERROR_VALUE;
    }
    case 'date': {
      // A date the text does not spell stays text.
      const serial = isoToSerial(String(value), date1904);
      return serial === undefined ? String(value) : serial;
    }
  }
  return null;
}

/** Whether a cell holds a value at all; a formula without one was never calculated. */
function hasValue(cell: Cell): boolean {
  return cell.kind !== 'blank' && cell.value !== null && cell.value !== undefined;
}

/** Whether what was calculated is what the file holds; numbers may differ in their last digits. */
function sameValue(stored: Scalar, value: Scalar): boolean {
  if (stored === value) return true;
  if (typeof stored === 'number' && typeof value === 'number') return Math.abs(stored - value) <= SAME_NUMBER * Math.max(Math.abs(stored), Math.abs(value));
  if (isError(stored) && isError(value)) return stored.error === value.error;
  return false;
}

function rangeKey(range: CellRange): string {
  return `${range.top}:${range.left}:${range.bottom}:${range.right}`;
}

/** First index whose key is not below `value`. */
function lowerBound(keys: number[], value: number): number {
  let low = 0;
  let high = keys.length;
  while (low < high) {
    const mid = (low + high) >>> 1;
    if (keys[mid] < value) low = mid + 1;
    else high = mid;
  }
  return low;
}

/** First position from `i` on that was not struck out; the way there is shortened for the next time. */
function find(skip: Int32Array, i: number): number {
  let root = i;
  while (skip[root] !== root) root = skip[root];
  while (skip[i] !== root) {
    const next = skip[i];
    skip[i] = root;
    i = next;
  }
  return root;
}

/** Whether `text` has `word` at `at`, without regard to case; `word` is in lower case. */
function wordAt(text: string, at: number, word: string): boolean {
  return text.substring(at, at + word.length).toLowerCase() === word;
}

/**
 * Whether a number format given to TEXT uses letters the format engine does not know as
 * placeholders: the day, month and year letters of other languages ("gg.aa.yyyy", "TT.MM.JJJJ").
 * Excel reads the format of TEXT in the language of the user; this engine reads en-US.
 */
function foreignFormat(code: string): boolean {
  const n = code.length;
  for (let i = 0; i < n; i++) {
    const c = code.charCodeAt(i);
    if (c === 34) {
      // "..."
      i++;
      while (i < n && code.charCodeAt(i) !== 34) i++;
    } else if (c === 92 || c === 95 || c === 42) i++;
    else if (c === 91) {
      // [...]
      while (i < n && code.charCodeAt(i) !== 93) i++;
    } else {
      const char = code.charAt(i);
      const lower = char.toLowerCase();
      // Not a letter.
      if (lower === char.toUpperCase()) continue;
      if (wordAt(code, i, 'general')) i += 6;
      else if (wordAt(code, i, 'am/pm')) i += 4;
      else if (wordAt(code, i, 'a/p')) i += 2;
      else if ('ymdhseb'.indexOf(lower) < 0) return true;
    }
  }
  return false;
}

/**
 * The names a formula binds with LET or as parameters of LAMBDA, in upper case. Such a name
 * is no defined name of the workbook where it is used; a use outside the call that binds it
 * is taken for one as well (it is #NAME? when calculated, as in Excel).
 */
function boundNames(root: Node): Set<string> | undefined {
  let found: Set<string> | undefined;
  const nodes: Node[] = [root];
  while (nodes.length > 0) {
    const node = nodes.pop() as Node;
    if (node.type === 'call') {
      if (BINDERS.has(node.name)) {
        const args = node.args;
        for (let k = 0; k < args.length - 1; k++) {
          const arg = args[k];
          // LET binds every other argument; LAMBDA every argument but its body.
          if (arg && arg.type === 'name' && arg.sheet === undefined && (node.name === 'LAMBDA' || k % 2 === 0)) (found || (found = new Set())).add(arg.name.toUpperCase());
        }
      }
      for (const arg of node.args) if (arg) nodes.push(arg);
    } else if (node.type === 'unary') nodes.push(node.operand);
    else if (node.type === 'binary') nodes.push(node.left, node.right);
  }
  return found;
}

/**
 * Whether a cell holds nothing a spill could cover up: no formula, no value, nothing else of the file (a shared
 * string, even an empty one, an inline string, value metadata, a phonetic mark); a format alone is fine.
 */
function isVacant(cell: Cell): boolean {
  return (
    !cell.formula &&
    cell.kind === 'blank' &&
    (cell.value === null || cell.value === undefined) &&
    cell.si === undefined &&
    cell.vm === undefined &&
    cell.cm === undefined &&
    !cell.ph &&
    !cell.runs
  );
}

/**
 * Rows an engine made (for the cells of a spill): such a row goes when the spill leaves it empty. A row of
 * the workbook stays, though it holds nothing (a file may have `<row r="5"/>`). Shared by all engines.
 */
const MADE_ROWS = new WeakSet<Row>();

/** A copy of a row with other cells; one an engine made stays one. */
function copyRow(row: Row, cells: Cell[]): Row {
  const next: Row = { ...row, cells };
  if (MADE_ROWS.has(row)) MADE_ROWS.add(next);
  return next;
}

const meets = (a: CellRange | undefined, b: CellRange) => !!a && a.top <= b.bottom && a.bottom >= b.top && a.left <= b.right && a.right >= b.left;

const unionOf = (a: CellRange, b: CellRange): CellRange => ({ top: Math.min(a.top, b.top), left: Math.min(a.left, b.left), bottom: Math.max(a.bottom, b.bottom), right: Math.max(a.right, b.right) });

const spillSize = (range: CellRange) => (range.bottom - range.top + 1) * (range.right - range.left + 1);

const sameRange = (a: CellRange | undefined, b: CellRange | undefined) =>
  a === b || (!!a && !!b && a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right);

/**
 * Where a dynamic array formula spills as the model says it: the range of its `ref`, which
 * starts at its own cell; undefined for a formula with none, or one that does not start there.
 */
function spillAreaOf(formula: CellFormula, row: number, col: number): CellRange | undefined {
  if (typeof formula.ref !== 'string') return undefined;
  const range = parseRange(formula.ref);
  if (!range || range.top !== row || range.left !== col) return undefined;
  return (range.bottom - range.top + 1) * (range.right - range.left + 1) <= MAX_SPILL_CELLS ? range : undefined;
}

/**
 * Cells whose value an engine wrote: such a value is not one a file vouches for. Cells are
 * never changed in place, so a cell stays in the set exactly as long as it holds that value.
 * Shared by all engines: a second engine over the same workbook knows what the first one wrote.
 */
const OWN = new WeakSet<Cell>();

export class CalcEngine {
  /**
   * Cells found in a circular reference by the last recalculation; when it took several
   * calls, by all of them. They hold 0, unless nothing they read changed and the file has
   * another value for them (a workbook with iterative calculation): then that value stays.
   */
  circular: CellAddress[] = [];
  /**
   * The last recalculate() call ran out of budget: the graph is not built yet, or some
   * formulas were not calculated and still hold their previous values. Call recalculate([])
   * (or with the next changes) to go on. Every call does part of the work that is left: a
   * workbook of N formulas is done after a number of calls proportional to N at most,
   * however small the budget.
   */
  incomplete = false;

  private readonly functions: Map<string, FunctionSpec>;
  private readonly context: MutableContext;
  private readonly evalOptions: EvaluateOptions;
  private sheets: SheetState[] = [];
  private sheetByName = new Map<string, number>();
  private names = new Map<string, { global?: NameEntry; local?: Map<number, NameEntry> }>();
  private volatiles = new Set<FormulaRec>();
  /** Formulas whose value depends on which rows are hidden. */
  private rowAware = new Set<FormulaRec>();
  /** Formulas that name their own cell. */
  private selfRefs = new Set<FormulaRec>();
  /** Formulas that use defined names. */
  private named = new Set<FormulaRec>();
  /** The names used by formulas listed in `unsupported`, by the key of the cell there. */
  private unsupportedNames = new Map<string, string[]>();
  /** Names met by supports(), for the formula being read. */
  private metNames: string[] = [];
  /** Sheets whose hidden rows changed since the formulas that care were last marked. */
  private rowSheets = new Set<number>();
  private unparsedCells = new Map<string, CellAddress>();
  private unsupportedCells = new Map<string, CellAddress>();
  private costlyCells = new Map<string, CellAddress>();
  private mismatchedCells = new Map<string, CellAddress>();
  /**
   * Formula cells that were read and left out of the graph (not parsed, not supported, a
   * shared formula without its master, a data table): a look through the model for formulas
   * the engine was not told about passes them by.
   */
  private looked = new WeakSet<Cell>();
  /** Cells whose value an engine wrote (see OWN). */
  private readonly own = OWN;
  /**
   * The formulas of the recalculation in progress, and how far it has come: everything in
   * `dirty` before `cursor` is done. A call that only goes on with it takes it up there.
   */
  private dirty: FormulaRec[] = [];
  private cursor = 0;
  /** The formula a call ran out of budget in: the next call starts with it. */
  private first: FormulaRec | undefined;
  /** Formulas finished by this call, to be written. */
  private done: FormulaRec[] = [];
  /** What evaluateText() keeps, apart from what a recalculation in progress keeps. */
  private previewMemos = new Map<string, unknown>();
  /** A preview is being calculated: every cell answers from the model. */
  private previewing = false;
  /** What functions keep for the rest of a recalculation (see CalcContext.memo). */
  private memos = new Map<string, unknown>();
  private work = 0;
  private workLimit = Infinity;
  private timeLimit = Infinity;
  private softWork = Infinity;
  private softTime = Infinity;
  private hardWork = Infinity;
  private hardTime = Infinity;
  private started = 0;
  /** Amount of work at which the budget is looked at next. */
  private nextCheck = Infinity;
  /** Formulas are being calculated: a budget that runs out before the first one is done is extended. */
  private calculating = false;
  private extended = false;
  /** Formulas finished by this call. */
  private finished = 0;
  /** The formula this call started with. */
  private sliceRoot: FormulaRec | undefined;
  private prepared = false;
  private built = false;
  private building: BuildState | undefined;
  private epoch = 0;
  private nowValue: number | undefined;
  /** Formulas whose dependents are still to be marked. */
  private walk: FormulaRec[] = [];
  /** Set while the dependents of formulas are only gathered, not marked (see `stale`). */
  private gathering: Set<FormulaRec> | undefined;
  private outdated: FormulaRec[] = [];
  /** Formulas being calculated, each one needed by the one before it. */
  private chain: FormulaRec[] = [];
  /** Length of the chain when the formula now on top of the JavaScript stack was started. */
  private floor = 0;
  private pending: FormulaRec | undefined;
  /** recalculate() was called for everything and that has not been done yet. */
  private wantAll = false;
  private wantOverwrite = false;
  /** Changed cells not looked at yet, and those whose dependents are not marked yet. */
  private wantCells: CellAddress[] = [];
  private ready: ReadyCell[] = [];
  // Circular references.
  private searchAll = false;
  private seeds: FormulaRec[] = [];
  private search: ComponentRun | undefined;
  private stamps = 0;
  private groupIds = 0;
  private groups = new Map<number, FormulaRec[]>();
  /** Formulas marked `looped`. */
  private loops: FormulaRec[] = [];
  private circularKeys = new Set<string>();
  private staleCells: CellAddress[] | undefined;
  /** Cells a dynamic array formula that went had filled and that were emptied: what reads them is calculated by the next recalculation. */
  private cleared: { sheet: number; range: CellRange }[] = [];
  /** Dynamic array formulas refused for the budget of all spills (see allocate()). */
  private budgetRefused = new Set<FormulaRec>();
  /** Dynamic array formulas that spill, or want to, into more than one cell: what allocate() shares the budget among. */
  private wanting = new Set<FormulaRec>();
  private allocatedEpoch = -1;
  /** While spills are being settled (see resolve()): the formulas whose spill is to change at the end of the round. */
  private deferred: Set<FormulaRec> | undefined;
  /** Cells of array formulas of the file that may still be made in this recalculation (see store()). */
  private fillLeft = MAX_ARRAY_FILL;
  /** Dynamic array formulas of the file that show #SPILL!: where they would spill is not known yet (see begin()). */
  private unguarded: FormulaRec[] = [];
  /** Those of `unguarded` calculated in this recalculation: a circle the file was saved with may run through them. */
  private reguarded: FormulaRec[] = [];
  /** Circularity found without anything being calculated for it (see rebuildForced()). */
  private quiet = false;
  /** Cells spills left since results were last written (see valueOf()). */
  private vacated: { sheet: number; range: CellRange }[] = [];
  /** Dynamic array formulas that were #SPILL! for being taken up too often (see recalculate()). */
  private tooOften = new Set<FormulaRec>();
  /** Formulas found to be circular, or no longer, since they were last calculated (see setCircular()). */
  private flipped: FormulaRec[] = [];
  private readonly spillLimit: number;
  private readonly bookLimit: number;
  private batching = false;
  private batchOpen = false;

  constructor(
    private readonly workbook: Workbook,
    private readonly options: CalcEngineOptions = {},
  ) {
    this.functions = options.functions || FUNCTIONS;
    const limits = options.limits || {};
    const limit = (value: number | undefined, most: number) => (typeof value === 'number' && value >= 1 ? Math.min(Math.floor(value), most) : most);
    this.spillLimit = limit(limits.spillCells, MAX_SPILL_CELLS);
    this.bookLimit = limit(limits.workbookSpillCells, MAX_WORKBOOK_SPILL_CELLS);
    const engine = this;
    this.context = {
      sheet: 0,
      row: 0,
      col: 0,
      get date1904() {
        return !!workbook.date1904;
      },
      get locale() {
        return options.locale ? options.locale() : undefined;
      },
      cellValue: (sheet, row, col) => engine.cellValue(sheet, row, col),
      forEachCell: (sheet, range, visit) => engine.forEachCell(sheet, range, visit),
      usedExtent: (sheet, range) => engine.contentExtent(sheet, range),
      sheetIndex: (name) => engine.sheetIndex(name),
      sheetName: (index) => (workbook.sheets[index] ? workbook.sheets[index].name : undefined),
      sheetCount: () => workbook.sheets.length,
      now: () => engine.now(),
      random: () => {
        const value = options.random ? options.random() : Math.random();
        return value >= 0 && value < 1 ? value : 0;
      },
      cellFormula: (sheet, row, col) => engine.cellFormula(sheet, row, col),
      memo: <T>(key: string, build: () => T): T => {
        const memos = engine.memos;
        if (memos.has(key)) return memos.get(key) as T;
        const value = build();
        memos.set(key, value);
        return value;
      },
      tick: (units) => engine.spend(units),
      rowHidden: (sheet, row) => engine.rowHidden(sheet, row, false),
      rowFiltered: (sheet, row) => engine.rowHidden(sheet, row, true),
      spillRange: (sheet, row, col) => engine.spillRangeOf(sheet, row, col),
    };
    this.evalOptions = {
      resolveName: (name, sheet) => engine.resolveName(name, sheet),
      cellFormula: (sheet, row, col) => engine.cellFormula(sheet, row, col),
      rowShift: 0,
      colShift: 0,
    };
  }

  /**
   * Formula cells whose text could not be parsed. Their values are left as the file has
   * them, so that syntax this engine does not know is not damaged.
   */
  get unparsed(): CellAddress[] {
    return Array.from(this.unparsedCells.values());
  }

  /**
   * Formula cells that parse but use what this engine cannot calculate: a function it does
   * not have, LET or LAMBDA, a reference across sheets (3D) or into another workbook, a sheet
   * or a name the workbook does not have, a defined name that uses one of those; or a use of
   * a function that is only partly supported: CELL with anything but "row", "col", "address"
   * or "contents", INDIRECT in R1C1 notation, INDEX with an area other than 1, a union of
   * references where no function takes it apart, TEXT with a format that has the date letters
   * of another language. They keep the values the file has for them, and formulas that read
   * them read those values. A formula without a value of the file (one just typed in, or one
   * whose value the engine wrote) is not listed: it is calculated and shows the error.
   */
  get unsupported(): CellAddress[] {
    return Array.from(this.unsupportedCells.values());
  }

  /**
   * Formula cells that could not be calculated even within the hard budget (see CalcBudget).
   * They keep the values they had. They are tried again, once, whenever something they read
   * changes, when everything is recalculated and after invalidate().
   */
  get overBudget(): CellAddress[] {
    return Array.from(this.costlyCells.values());
  }

  /**
   * The cells of `overBudget` and every formula that reads one of them, directly or through
   * other formulas: what they show may not fit the cells they are calculated from.
   */
  get stale(): CellAddress[] {
    if (!this.staleCells) {
      const found = new Set<FormulaRec>();
      this.costlyCells.forEach((cell) => {
        const sh = this.sheets[cell.sheet];
        const rec = sh ? sh.recs.get(keyOf(cell.row, cell.col)) : undefined;
        if (rec) found.add(rec);
      });
      const walk = Array.from(found);
      this.gathering = found;
      this.walk = walk;
      for (let k = 0; k < walk.length; k++) this.dependents(walk[k], false);
      this.gathering = undefined;
      this.walk = [];
      this.staleCells = Array.from(found).map((rec) => ({ sheet: rec.sheet, row: rec.row, col: rec.col }));
    }
    return this.staleCells.slice();
  }

  /**
   * Formula cells that were calculated although nothing they read had changed (everything
   * was recalculated, or the cell was named as changed with the formula it had) and came out
   * different from the value the file has for them. That value stays, and formulas that read
   * the cell read it: the program that wrote the file knew something this engine does not
   * (a number format in another language, say). A cell leaves the list when something it
   * reads changes: from then on its value is the engine's.
   */
  get mismatched(): CellAddress[] {
    return Array.from(this.mismatchedCells.values());
  }

  /**
   * Formula cells that hold no value: a file written by a program that does not calculate
   * stores its formulas that way. recalculate(engine.uncalculated()) gives them their values.
   * Cells the engine cannot calculate are not listed once it has looked at them (see
   * `unparsed`); a formula written since the engine read the workbook is listed, whether or
   * not the engine was told about it.
   */
  uncalculated(): CellAddress[] {
    const out: CellAddress[] = [];
    const sheets = this.workbook.sheets || [];
    for (let s = 0; s < sheets.length; s++) {
      const sh = this.built ? this.sheets[s] : undefined;
      sheets[s].rows.forEach((row, index) => {
        if (!isIndex(index, MAX_ROWS) || !row || !row.cells) return;
        row.cells.forEach((cell, col) => {
          if (!cell || !cell.formula || hasValue(cell) || col >= MAX_COLS || cell.formula.type === 'dataTable') return;
          if (sh) {
            const rec = sh.recs.get(keyOf(index, col));
            // Over budget: it was tried. Read and left out of the graph: it cannot be parsed, or its
            // master is missing. Not in the graph and never read: written since, and waiting.
            if (rec ? rec.formula === cell.formula && rec.costly : this.looked.has(cell) && !this.wantCells.some((want) => want.sheet === s && want.row === index && want.col === col)) return;
          }
          out.push({ sheet: s, row: index, col });
        });
      });
    }
    return out;
  }

  /**
   * Recalculates what depends on the given cells (all formulas when `changed` is omitted:
   * the model is then looked through for formulas that were written, replaced or removed
   * without the engine being told, and those are calculated like cells that were named,
   * with or without `overwrite`).
   * Writes results into the model by replacing cell objects (never mutating them) and
   * marks sheet.changedRows / sheet.modified for cells whose value changed.
   * Returns the cells whose value changed, per sheet.
   *
   * `changed` has to name every cell whose content or formula was changed since the last
   * call. Never throws: a formula that cannot be calculated gets an error value.
   *
   * The list of sheets is looked at by every call: when it is not the one the engine read
   * (a sheet inserted, deleted, moved or renamed), the graph is forgotten as by invalidate()
   * and built again; when a sheet went that formulas of the graph read, or a calculation was
   * under way, everything is calculated with `overwrite` as well (and `incomplete` is set until
   * it is): the formulas that read the sheet were rewritten, and the engine was not told which
   * they are. Calling invalidate() first, as the view does, keeps the choice of what to
   * calculate with the caller.
   *
   * The call stops when its budget (options.budget) is used up and sets `incomplete`. What
   * was calculated by then is written; the other formulas keep their previous values, never
   * a half-calculated one, and are taken up by the next call, whatever its `changed` is.
   * Nothing is written before the dependency graph is complete. A formula finished by one
   * call is not calculated again by the calls that go on with the rest.
   */
  recalculate(changed?: CellAddress[], options?: RecalculateOptions): { sheet: number; cells: { row: number; col: number }[] }[] {
    this.prepare(options !== undefined && options.sheetsOfBefore === true);
    this.batchOpen = false;
    this.startBudget();
    for (const sh of this.sheets) {
      sh.copied.clear();
      sh.changed = [];
    }
    if (!Array.isArray(changed)) {
      this.wantAll = true;
      if (options && options.overwrite === true) this.wantOverwrite = true;
      // Not told which cells changed: how far the columns hold something is found again (see contentRows()).
      for (const sh of this.sheets) {
        sh.lastRows = undefined;
        sh.contentVersion++;
        sh.rowKeys = undefined;
      }
      // A graph that was built before may have missed what was written since.
      if (this.built) this.findUnknown();
    } else {
      for (const cell of changed) {
        if (!cell || !this.sheets[cell.sheet] || !isIndex(cell.row, MAX_ROWS) || !isIndex(cell.col, MAX_COLS)) continue;
        this.wantCells.push({ sheet: cell.sheet, row: cell.row, col: cell.col });
      }
    }
    this.incomplete = true;
    if (!this.build()) return [];
    // The graph follows the cells that were changed; only then can circular references be looked for.
    const strays: CellAddress[] = [];
    for (const cell of this.wantCells) {
      const sh = this.sheets[cell.sheet];
      const modelRow = sh.model.rows.get(cell.row);
      const held = modelRow ? modelRow.cells[cell.col] : undefined;
      if (held && held.spilled === true && !held.formula) strays.push(cell);
      this.noteContent(sh, cell.row, cell.col);
      const before = sh.recs.get(keyOf(cell.row, cell.col));
      this.refresh(sh, cell.row, cell.col);
      const after = sh.recs.get(keyOf(cell.row, cell.col));
      this.ready.push({ sheet: cell.sheet, row: cell.row, col: cell.col, replaced: after !== undefined && after !== before });
    }
    this.wantCells = [];
    // A call that goes on with a calculation cut short leaves the circles its spills made to the end of it (see
    // settleSpills()): looked for now, they would take formulas up again with every slice.
    const goingOn = this.ready.length === 0 && !this.wantAll && (this.first !== undefined || this.cursor < this.dirty.length);
    if (!goingOn && !this.components()) return [];

    const all = this.wantAll;
    const changes = this.ready.length > 0;
    const pending = this.first !== undefined || this.cursor < this.dirty.length;
    // Marking every formula takes a while in a large workbook: not at the end of a call that built the graph.
    if (all && this.outOfTime()) return [];
    // Rows were hidden or shown: a change for the formulas that leave hidden rows out.
    const rowReaders = this.rowReaders();
    // A call that only goes on with what was left takes it up where the last one stopped: what is
    // done stays done, and what functions keep for a recalculation is still good.
    if (all || changes || rowReaders.length > 0 || !pending) this.begin(all, changes, pending, rowReaders);
    if (!goingOn) this.retake(this.takeFlipped());
    this.calculateAll();
    this.flush();
    // A spill that came or went makes or breaks circles (a formula reading the cells another one
    // spills into, which reads it): they are looked for now, and what they change is calculated
    // again, so that a recalculation ends where one of a graph built afresh does.
    // A formula taken up too often (see MAX_REOPENS) is #SPILL! for now; it is tried again the same way, with its count
    // started over, so that what came out of a circle found since, or of the order formulas came in, does not stay.
    // The circles a spill made or broke, and the spills that would not settle, are worked out once all else is
    // calculated: the same way whatever the order and the slices formulas were calculated in (see resolve()).
    if (this.first === undefined) this.settleSpills();
    if (this.first === undefined && this.reguarded.length > 0) this.rebuildForced();
    // A changed cell marked as a spill's that no spill fills (an undo put it back after the spill went) is emptied.
    if (this.first === undefined) for (const cell of strays) this.dropStray(this.sheets[cell.sheet], cell.row, cell.col);
    // Formulas that turned out to have been changed behind the engine's back are read again for next time.
    for (const rec of this.outdated) this.refresh(this.sheets[rec.sheet], rec.row, rec.col);
    this.outdated = [];
    const dirty = this.dirty;
    while (this.cursor < dirty.length && (dirty[this.cursor].state === DONE || dirty[this.cursor].epoch !== this.epoch)) this.cursor++;
    this.incomplete = this.first !== undefined || this.cursor < dirty.length || this.search !== undefined || this.flipped.length > 0;
    if (!this.incomplete) {
      this.dirty = [];
      this.cursor = 0;
      this.memos.clear();
    }

    const out: { sheet: number; cells: { row: number; col: number }[] }[] = [];
    for (const sh of this.sheets) if (sh.changed.length > 0) out.push({ sheet: sh.index, cells: sh.changed });
    return out;
  }

  /** Starts a recalculation: marks what has to be calculated. */
  private begin(all: boolean, changes: boolean, pending: boolean, rowReaders: FormulaRec[]): void {
    // What the last call left over comes first. Its dependents were marked when it was: the ones
    // that are done did not read it.
    const carried: FormulaRec[] = [];
    if (pending) {
      const epoch = this.epoch;
      if (this.first) carried.push(this.first);
      for (let k = this.cursor; k < this.dirty.length; k++) {
        const rec = this.dirty[k];
        if (rec.state !== DONE && rec.epoch === epoch && rec !== this.first) carried.push(rec);
      }
    }
    this.epoch++;
    this.nowValue = undefined;
    this.memos.clear();
    this.dirty = [];
    this.cursor = 0;
    this.first = undefined;
    this.walk = [];
    // Rows the sheet gained or lost since its row list was made (an edit adds them, the engine keeps its own up).
    for (const sh of this.sheets) if (sh.rowKeys && (sh.rowKeysOf !== sh.model.rows || sh.rowKeys.length !== sh.model.rows.size)) sh.rowKeys = undefined;
    // A new recalculation: what earlier ones found while calculating may no longer hold.
    this.circular = [];
    this.circularKeys.clear();
    for (const rec of this.loops) rec.looped = false;
    this.loops = [];
    for (const rec of carried) if (!rec.gone) this.touch(rec);
    if (all) {
      const overwrite = this.wantOverwrite;
      let vouched = 0;
      for (const sh of this.sheets) {
        sh.recs.forEach((rec) => {
          this.touch(rec);
          if (overwrite || rec.live) rec.live = true;
          else if (rec.timed || this.unvalued(rec)) {
            // Its value changes by itself, or it has none: whatever reads it is calculated anew.
            rec.live = true;
            this.walk.push(rec);
          } else if (this.storedOf(rec) !== undefined) vouched++;
        });
      }
      // Without a value of the file to keep, nothing needs to know what is live.
      if (vouched === 0) this.walk = [];
    }
    for (const cell of this.ready) {
      const sh = this.sheets[cell.sheet];
      const rec = sh.recs.get(keyOf(cell.row, cell.col));
      if (rec) this.mark(rec, cell.replaced);
      else this.collect(sh, cell.row, cell.col, cell.row, cell.col, true);
    }
    for (const area of this.cleared) {
      const sh = this.sheets[area.sheet];
      if (sh) this.collect(sh, area.range.top, area.range.left, area.range.bottom, area.range.right, true);
    }
    this.cleared = [];
    // Which rows are hidden is something these formulas read: their results replace the values of the file.
    for (const rec of rowReaders) if (!rec.gone) this.mark(rec, true);
    for (const rec of this.unguarded) if (!rec.gone) this.mark(rec, false);
    for (const rec of this.unguarded) this.reguarded.push(rec);
    this.unguarded = [];
    // A whole column or row reads as far as its sheet holds something: when that changed (a value typed or cleared
    // anywhere on it), what reads one is calculated again.
    // Looked at once per whole column or row read, not per formula: most read the same few.
    for (const sh of this.sheets) {
      if (sh.wideRanges.size === 0) continue;
      sh.wideRanges.forEach((wide) => {
        const range = wide.range;
        const extent = this.contentExtent(sh.index, range);
        const now = `${range.bottom > range.top ? extent.rows : ''}:${range.right > range.left ? extent.cols : ''}`;
        const before = wide.seen;
        wide.seen = now;
        if (before !== undefined && before !== now) wide.readers.forEach((rec) => {
          if (!rec.gone) this.mark(rec, true);
        });
      });
    }
    // What was #SPILL! for being taken up too often is tried again: the order of before need not come again.
    this.tooOften.forEach((rec) => {
      if (!rec.gone) this.mark(rec, true);
    });
    this.tooOften.clear();
    if (changes || !all) {
      this.volatiles.forEach((rec) => {
        // A formula that reads cells it chooses while calculating may read one that changed.
        if (!rec.costly) this.mark(rec, rec.timed || changes);
      });
    }
    this.wantAll = false;
    this.wantOverwrite = false;
    this.ready = [];
    // Everything that reads a dirty formula is dirty as well; `walk` grows while it is walked.
    const walk = this.walk;
    for (let k = 0; k < walk.length; k++) this.dependents(walk[k], walk[k].live);
    this.walk = [];
    this.fillLeft = MAX_ARRAY_FILL;
  }

  /**
   * Formula cells of the model that the graph does not have as they are: written, replaced
   * or removed since the graph was built without the engine being told. They are taken as
   * cells named to recalculate(). One pass over the cells of the workbook.
   */
  private findUnknown(): void {
    const want = this.wantCells;
    const looked = this.looked;
    for (const sh of this.sheets) {
      const recs = sh.recs;
      const s = sh.index;
      let known = 0;
      sh.model.rows.forEach((row, index) => {
        if (!isIndex(index, MAX_ROWS) || !row || !row.cells) return;
        const cells = row.cells;
        const length = cells.length < MAX_COLS ? cells.length : MAX_COLS;
        for (let col = 0; col < length; col++) {
          const cell = cells[col];
          if (!cell || !cell.formula) continue;
          const rec = recs.get(keyOf(index, col));
          if (rec && rec.formula === cell.formula) known++;
          else if (rec || !looked.has(cell)) want.push({ sheet: s, row: index, col });
        }
      });
      if (known === recs.size) continue;
      // Formulas of the graph whose cells are gone or hold something else.
      recs.forEach((rec) => {
        const row = sh.model.rows.get(rec.row);
        const cell = row && row.cells ? row.cells[rec.col] : undefined;
        if (!cell || !cell.formula) want.push({ sheet: s, row: rec.row, col: rec.col });
      });
    }
  }

  /**
   * Call after rows or columns were inserted or deleted, sheets changed, or defined names
   * changed: forgets the dependency graph. (A list of sheets that changed is noticed without
   * it as well: see recalculate().) The recalculation that follows builds it again,
   * in slices when it is large. When the edit can have changed what formulas are worth
   * (rows deleted, a name redefined), recalculate everything with `overwrite`.
   */
  invalidate(): void {
    this.prepared = false;
    this.built = false;
    this.building = undefined;
    this.sheets = [];
    this.sheetByName = new Map();
    this.names = new Map();
    this.volatiles = new Set();
    this.rowAware = new Set();
    this.selfRefs = new Set();
    this.named = new Set();
    this.unsupportedNames = new Map();
    this.unparsedCells = new Map();
    this.unsupportedCells = new Map();
    this.costlyCells = new Map();
    this.mismatchedCells = new Map();
    this.looked = new WeakSet();
    this.staleCells = undefined;
    // Whatever was left over is calculated with everything else: the caller recalculates all after this.
    this.incomplete = false;
    this.memos.clear();
    this.dirty = [];
    this.cursor = 0;
    this.first = undefined;
    this.done = [];
    this.outdated = [];
    this.chain = [];
    this.walk = [];
    this.wantAll = false;
    this.wantOverwrite = false;
    this.wantCells = [];
    this.ready = [];
    this.searchAll = false;
    this.seeds = [];
    this.search = undefined;
    this.groups = new Map();
    this.loops = [];
    this.cleared = [];
    this.budgetRefused = new Set();
    this.wanting = new Set();
    this.unguarded = [];
    this.vacated = [];
    this.flipped = [];
    this.tooOften = new Set();
    this.batchOpen = false;
  }

  /**
   * The value a formula text would have in a cell, without storing anything (for a formula bar preview).
   * Other formulas are not recalculated: they answer with the values the model holds.
   * Throws FormulaParseError when the text is not a formula, so that the caller can show
   * where it goes wrong; a formula that parses never throws.
   */
  evaluateText(formula: string, at: CellAddress): Scalar {
    return this.preview(parseFormula(formula), at, at, 0, 0);
  }

  /**
   * The value the formula of the cell `anchor` has when it is applied to the cell `at` of
   * the same sheet: its relative references move by the distance between the two, as they
   * do when a cell is copied (the formula of a conditional format, of a data validation).
   * The text is parsed once, however many cells it is calculated for. A reference that is
   * pushed off the sheet is #REF!. Throws FormulaParseError like evaluateText().
   */
  evaluateShifted(formula: string, anchor: CellAddress, at: { row: number; col: number }): Scalar {
    const node = parseFormula(formula);
    const row = at && isIndex(at.row, MAX_ROWS) ? at.row : 0;
    const col = at && isIndex(at.col, MAX_COLS) ? at.col : 0;
    const fromRow = anchor && isIndex(anchor.row, MAX_ROWS) ? anchor.row : 0;
    const fromCol = anchor && isIndex(anchor.col, MAX_COLS) ? anchor.col : 0;
    return this.preview(node, anchor, { row, col }, row - fromRow, col - fromCol);
  }

  /**
   * Runs `run` with everything evaluateText() and evaluateShifted() work out inside it shared
   * between the calls: one budget for all of them, and what functions keep for a
   * recalculation (the index of a lookup table) is kept from one call to the next. The
   * model must not change while `run` runs.
   */
  batch<T>(run: () => T): T {
    if (this.batching) return run();
    this.batching = true;
    this.batchOpen = false;
    try {
      return run();
    } finally {
      this.batching = false;
      this.batchOpen = false;
      this.previewMemos = new Map();
    }
  }

  private preview(node: Node, sheetOf: CellAddress, at: { row: number; col: number }, rowShift: number, colShift: number): Scalar {
    const context = this.context;
    const value = this.previewed<Scalar>(sheetOf, at, rowShift, colShift, ERROR_VALUE, () => toCellValue(evaluate(node, context, this.functions, this.evalOptions), context));
    if (typeof value === 'number' && !isFinite(value)) return ERROR_NUM;
    return value === null || value === undefined ? 0 : value;
  }

  /**
   * Runs `run` as a preview for the cell `at` of the sheet of `sheetOf`: no formula is
   * calculated, every cell answers from the model. `failed` is the answer when `run` throws
   * (out of budget, among other things).
   */
  private previewed<T>(sheetOf: CellAddress, at: { row: number; col: number }, rowShift: number, colShift: number, failed: T, run: () => T): T {
    this.prepare();
    if (!this.batchOpen) {
      this.nowValue = undefined;
      this.previewMemos = new Map();
      this.startBudget();
      for (const sh of this.sheets) sh.rowKeys = undefined;
      this.batchOpen = this.batching;
    }
    // What a recalculation in progress keeps is left alone.
    const memos = this.memos;
    this.memos = this.previewMemos;
    this.previewing = true;
    const context = this.context;
    context.sheet = sheetOf && isIndex(sheetOf.sheet, this.sheets.length) ? sheetOf.sheet : 0;
    context.row = at && isIndex(at.row, MAX_ROWS) ? at.row : 0;
    context.col = at && isIndex(at.col, MAX_COLS) ? at.col : 0;
    this.evalOptions.rowShift = rowShift;
    this.evalOptions.colShift = colShift;
    this.evalOptions.arrayFormula = false;
    this.evalOptions.dynamic = false;
    let value: T;
    try {
      value = run();
    } catch (error) {
      value = failed;
    }
    this.previewing = false;
    this.memos = memos;
    this.evalOptions.rowShift = 0;
    this.evalOptions.colShift = 0;
    if (!this.batchOpen) this.previewMemos = new Map();
    return value;
  }

  /**
   * Every value of what a formula text gives in a cell, in one call: the cells of a reference
   * (a range, a range of another sheet, a defined name, OFFSET, INDIRECT ...), the elements
   * of an array, or the one value of anything else. Like evaluateText() nothing is stored and
   * other formulas answer with the values the model holds; inside batch() it shares the budget
   * and what is kept with the other calls.
   *
   * A reference is cut to the used part of its sheet. No more than `maxCells` values are
   * given: a larger result comes with `truncated` set and its first values, row by row.
   * Never throws: a text that is not a formula and a formula that cannot be calculated within
   * the budget give #VALUE!, a formula whose value is an error gives that error. An error
   * inside a reference or an array is one of the values.
   */
  evaluateRange(formula: string, at: CellAddress, options?: EvaluateRangeOptions): RangeValues | CalcError {
    let node: Node;
    try {
      if (typeof formula !== 'string') return ERROR_VALUE;
      node = parseFormula(formula);
    } catch (error) {
      return ERROR_VALUE;
    }
    const asked = options ? Number(options.maxCells) : NaN;
    const limit = asked >= 1 ? Math.min(Math.floor(asked), MAX_RANGE_VALUES) : asked === asked ? 1 : DEFAULT_RANGE_VALUES;
    const row = at && isIndex(at.row, MAX_ROWS) ? at.row : 0;
    const col = at && isIndex(at.col, MAX_COLS) ? at.col : 0;
    const anchor = options ? options.anchor : undefined;
    const fromRow = anchor && isIndex(anchor.row, MAX_ROWS) ? anchor.row : row;
    const fromCol = anchor && isIndex(anchor.col, MAX_COLS) ? anchor.col : col;
    const context = this.context;
    return this.previewed<RangeValues | CalcError>(at, { row, col }, anchor ? row - fromRow : 0, anchor ? col - fromCol : 0, ERROR_VALUE, () =>
      this.valuesOf(evaluate(node, context, this.functions, this.evalOptions), limit),
    );
  }

  /** The values of what a formula gave, `limit` at most. */
  private valuesOf(value: Value, limit: number): RangeValues | CalcError {
    if (isRef(value)) {
      const range = value.range;
      const sh = this.sheets[value.sheet];
      if (!sh || !range || !isIndex(range.top, MAX_ROWS) || !isIndex(range.left, MAX_COLS) || !isIndex(range.bottom, MAX_ROWS) || !isIndex(range.right, MAX_COLS)) return calcError('#REF!');
      if (range.bottom < range.top || range.right < range.left) return calcError('#REF!');
      // Cut to the rows and columns the sheet has (this is shown, a list of choices, never written); at least its first cell stays.
      const extent = this.contentExtent(value.sheet, range);
      const bottom = Math.min(range.bottom, Math.max(extent.rows - 1, range.top));
      const right = Math.min(range.right, Math.max(extent.cols - 1, range.left));
      const rows = bottom - range.top + 1;
      const cols = right - range.left + 1;
      const total = rows * cols;
      const count = Math.min(total, limit);
      const values: Scalar[] = new Array(count).fill(null);
      const top = range.top;
      const left = range.left;
      this.forEachCell(value.sheet, { top, left, bottom: top + Math.ceil(count / cols) - 1, right }, (cell, r, c) => {
        const k = (r - top) * cols + (c - left);
        if (k >= 0 && k < count) values[k] = cell;
      });
      return { rows, cols, values, truncated: count < total, sheet: value.sheet, range: { top, left, bottom, right } };
    }
    if (isArray(value)) {
      const total = value.data.length;
      const count = Math.min(total, limit);
      const values: Scalar[] = new Array(count);
      for (let k = 0; k < count; k++) {
        const item = value.data[k];
        values[k] = item === undefined ? null : item;
      }
      this.spend(count);
      return { rows: value.rows, cols: value.cols, values, truncated: count < total };
    }
    if (isError(value)) return value;
    if (typeof value === 'number' && !isFinite(value)) return ERROR_NUM;
    return { rows: 1, cols: 1, values: [value === undefined ? null : value], truncated: false };
  }

  /**
   * Call after rows of a sheet were hidden or shown (by hand, by a filter); without a sheet,
   * for every sheet. The formulas that leave hidden rows out (SUBTOTAL, AGGREGATE) and read
   * that sheet are calculated by the next recalculate() call, and what they give replaces the
   * values of the file: which rows are hidden is something they read. Costs nothing by itself.
   */
  rowsChanged(sheet?: number): void {
    const count = (this.workbook.sheets || []).length;
    if (sheet === undefined) for (let s = 0; s < count; s++) this.rowSheets.add(s);
    else if (isIndex(sheet, count)) this.rowSheets.add(sheet);
  }

  /** The formulas rowsChanged() was called for; the graph has to be complete. */
  private rowReaders(): FormulaRec[] {
    const sheets = this.rowSheets;
    const out: FormulaRec[] = [];
    if (sheets.size === 0) return out;
    this.rowAware.forEach((rec) => {
      // A formula that chooses its cells while calculating may read any sheet.
      let reads = rec.volatile;
      for (let k = 0; !reads && k < rec.precedents.length; k++) reads = !rec.precedents[k].shapeOnly && sheets.has(rec.precedents[k].sheet);
      if (reads) out.push(rec);
    });
    sheets.clear();
    return out;
  }

  /**
   * Whether the dependency graph is complete and the search for circular references in it
   * done: cycles() answers only then. recalculate() works on both, in slices when the
   * workbook is large (see `incomplete`).
   */
  get graphReady(): boolean {
    return this.built && this.search === undefined && !this.searchAll && this.seeds.length === 0 && this.wantCells.length === 0;
  }

  /**
   * Every formula cell that is part of a circular reference made of the references the
   * formulas name, whether or not it was calculated yet (`circular` lists what recalculations
   * met, which includes circles through OFFSET and INDIRECT). Empty while `graphReady` is
   * false: a list of some of them would look like the whole.
   */
  cycles(): CellAddress[] {
    const out: CellAddress[] = [];
    if (!this.graphReady) return out;
    const seen = new Set<FormulaRec>();
    const add = (rec: FormulaRec) => {
      if (rec.gone || !rec.circular || seen.has(rec)) return;
      seen.add(rec);
      out.push({ sheet: rec.sheet, row: rec.row, col: rec.col });
    };
    this.selfRefs.forEach(add);
    this.groups.forEach((members) => members.forEach(add));
    return out;
  }

  /**
   * Call after workbook.definedNames changed (a name added, removed or given another
   * formula), instead of invalidate(): the names are read again, and the formulas that use a
   * name that is no longer what it was, directly or through other names, are read into the
   * graph again and calculated by the next recalculate() call, as formulas that changed.
   * Formulas that use no such name are left alone. Sheets, rows and columns must not have
   * changed.
   */
  namesChanged(): void {
    if (!this.prepared) return;
    if (!this.built) {
      // The formulas read so far were read with the old names: reading starts over. Nothing was calculated yet.
      const wantAll = this.wantAll;
      const wantOverwrite = this.wantOverwrite;
      const wantCells = this.wantCells;
      const incomplete = this.incomplete;
      this.invalidate();
      this.wantAll = wantAll;
      this.wantOverwrite = wantOverwrite;
      this.wantCells = wantCells;
      this.incomplete = incomplete;
      return;
    }
    const changed = this.readNames(this.names);
    if (changed.size === 0) return;
    this.previewMemos = new Map();
    const uses = (names: string[] | undefined): boolean => {
      for (let k = 0; names && k < names.length; k++) if (changed.has(names[k])) return true;
      return false;
    };
    const again: CellAddress[] = [];
    this.named.forEach((rec) => {
      if (uses(rec.names)) again.push({ sheet: rec.sheet, row: rec.row, col: rec.col });
    });
    this.unsupportedNames.forEach((names, key) => {
      const cell = this.unsupportedCells.get(key);
      if (cell && uses(names)) again.push(cell);
    });
    for (const cell of again) {
      this.reread(this.sheets[cell.sheet], cell.row, cell.col);
      this.ready.push({ sheet: cell.sheet, row: cell.row, col: cell.col, replaced: true });
    }
  }

  // -------------------------------------------------------------------------
  // What the evaluator asks for

  private now(): number {
    if (this.nowValue === undefined) {
      let serial: number;
      if (this.options.now) serial = Number(this.options.now());
      else {
        // Wall-clock time, as every other date in a sheet.
        const date = new Date();
        serial = (date.getTime() - date.getTimezoneOffset() * 60000) / 86400000 + 25569;
        if (this.workbook.date1904) serial -= 1462;
      }
      this.nowValue = isFinite(serial) ? serial : 0;
    }
    return this.nowValue;
  }

  private sheetIndex(name: string): number | undefined {
    return this.sheetByName.get(foldName(name));
  }

  /** The tree of a defined name as seen from a sheet: undefined when there is no such name, null when its formula cannot be parsed. */
  private nameTree(name: string, sheet: number): Node | null | undefined {
    const found = this.names.get(String(name).toUpperCase());
    if (!found) return undefined;
    const entry = (found.local && found.local.get(sheet)) || found.global;
    if (!entry) return undefined;
    if (entry.node === undefined) {
      try {
        entry.node = parseFormula(entry.text);
      } catch (error) {
        entry.node = null;
      }
    }
    return entry.node;
  }

  private resolveName(name: string, sheet: number): Node | undefined {
    return this.nameTree(name, sheet) || undefined;
  }

  // -------------------------------------------------------------------------
  // The budget
  //
  // A call has a budget, and formulas are given up to the next call when it runs out. But a
  // formula that needs more than a whole budget would never be done that way. So when the
  // budget runs out before the call has finished a single formula, the call goes on up to
  // the hard limit, until that formula is done. Only a formula that had the hard budget to
  // itself and did not make it is given up for good.

  private startBudget(): void {
    const budget = this.options.budget;
    const milliseconds = budget && budget.milliseconds !== undefined ? Number(budget.milliseconds) : DEFAULT_MILLISECONDS;
    const work = budget && budget.work !== undefined ? Number(budget.work) : Infinity;
    this.softTime = milliseconds > 0 ? milliseconds : Infinity;
    this.softWork = work >= 0 ? work : Infinity;
    const hardTime = budget && budget.hardMilliseconds !== undefined ? Number(budget.hardMilliseconds) : NaN;
    const hardWork = budget && budget.hardWork !== undefined ? Number(budget.hardWork) : NaN;
    this.hardTime = hardTime === 0 ? Infinity : hardTime > 0 ? Math.max(hardTime, this.softTime) : Math.max(this.softTime * HARD_FACTOR, MIN_HARD_MILLISECONDS);
    this.hardWork = hardWork >= 0 ? Math.max(hardWork, this.softWork) : Math.max(this.softWork * HARD_FACTOR, MIN_HARD_WORK);
    if (this.softTime === Infinity) this.hardTime = Infinity;
    this.timeLimit = this.softTime;
    this.workLimit = this.softWork;
    this.work = 0;
    this.extended = false;
    this.finished = 0;
    this.sliceRoot = undefined;
    this.calculating = false;
    this.started = this.timeLimit < Infinity ? this.clock() : 0;
    this.nextCheck = Math.min(CLOCK_EVERY, this.workLimit + 1);
  }

  private clock(): number {
    const value = this.options.clock ? Number(this.options.clock()) : Date.now();
    return value === value ? value : 0;
  }

  /** Whether the time of this call is used up; for the work that is not calculating. */
  private outOfTime(): boolean {
    return this.timeLimit < Infinity && this.clock() - this.started > this.timeLimit;
  }

  /** Counts work; gives the calculation up when the budget is used up. */
  private spend(units: number): void {
    this.work += units;
    if (this.work < this.nextCheck) return;
    for (;;) {
      let over = this.work > this.workLimit;
      if (!over) {
        this.nextCheck = Math.min(this.work + CLOCK_EVERY, this.workLimit + 1);
        over = this.timeLimit < Infinity && this.clock() - this.started > this.timeLimit;
      }
      if (!over) return;
      if (!this.calculating || this.finished > 0 || this.extended) throw BUDGET;
      // Not one formula done: this call goes on until one is, within the hard limit.
      this.extended = true;
      this.workLimit = this.hardWork;
      this.timeLimit = this.hardTime;
    }
  }

  private cellFormula(sheet: number, row: number, col: number): string | undefined {
    const sh = this.sheets[sheet];
    const modelRow = sh ? sh.model.rows.get(row) : undefined;
    const cell = modelRow ? modelRow.cells[col] : undefined;
    if (!cell || !cell.formula) return undefined;
    if (cell.formula.text !== '') return cell.formula.text;
    const rec = sh.recs.get(keyOf(row, col));
    if (rec && rec.formula === cell.formula) return shiftFormula(rec.text, rec.dRow, rec.dCol);
    // A cell of a shared formula the graph does not know (it was not built, or the master is missing).
    if (cell.formula.type === 'shared' && cell.formula.si !== undefined) {
      const master = sharedMasters(sh.model).get(cell.formula.si);
      if (master) return shiftFormula(master.text, row - master.row, col - master.col);
    }
    return '';
  }

  /** Whether a row is hidden; with `filtered`, only when a filter with a criterion has it among its data rows. */
  private rowHidden(sheet: number, row: number, filtered: boolean): boolean {
    const sh = this.sheets[sheet];
    const modelRow = sh ? sh.model.rows.get(row) : undefined;
    if (!modelRow || modelRow.hidden !== true) return false;
    if (!filtered) return true;
    // The model has one flag for both ways of hiding. In a filtered list Excel leaves out
    // every hidden row, so a hidden row under the header of a filter in use counts as filtered.
    const filter = sh.model.autoFilter;
    if (!filter || !filter.range || !(row > filter.range.top && row <= filter.range.bottom) || !Array.isArray(filter.columns)) return false;
    for (const column of filter.columns) if (column && column.kind !== 'none') return true;
    return false;
  }

  /**
   * One past the last row and column of a sheet that hold something (CalcContext.usedExtent): a value or a
   * formula anywhere on the sheet, and the spills (written or not, of any array formula) that reach into
   * `range`, the range being read. It comes from what the sheet holds, never from how far it was written
   * once (Sheet.rowCount): two people who did the same edits in other orders, or one who opened the saved
   * file, get the same. A spill elsewhere does not count: "=A:A*2" would otherwise read further as its own
   * spill grows. The part from values and formulas is worked out again after each change of the workbook.
   */
  private contentExtent(sheet: number, range?: CellRange): { rows: number; cols: number } {
    const sh = this.sheets[sheet];
    if (!sh) return { rows: 0, cols: 0 };
    const engine = this;
    // Each side worked out only when asked for: a whole column needs the last row, a whole row the last column.
    return {
      get rows() {
        let rows = engine.contentRows(sh, range);
        if (range) engine.reachOf(sh, range, (spill) => {
          if (spill.bottom + 1 > rows) rows = spill.bottom + 1;
        });
        return rows;
      },
      get cols() {
        let cols = engine.contentCols(sh, range);
        if (range) engine.reachOf(sh, range, (spill) => {
          if (spill.right + 1 > cols) cols = spill.right + 1;
        });
        return cols;
      },
    };
  }

  /** Whether a cell holds something for contentExtent(): a formula, or a value that is not a spill's. */
  private static holds(cell: Cell | undefined): boolean {
    return !!cell && (!!cell.formula || (cell.spilled !== true && cell.value !== null && cell.value !== undefined));
  }

  /**
   * One past the last row that holds something in the columns of `range` (all columns without one). A whole column
   * reads as far as ITS columns hold something, so that what is typed elsewhere on the sheet changes nothing of it.
   */
  private contentRows(sh: SheetState, range?: CellRange): number {
    const last = this.lastRows(sh);
    const left = range ? range.left : 0;
    const right = range ? range.right : MAX_COLS - 1;
    let rows = 0;
    if (right - left < 256) {
      for (let c = left; c <= right; c++) rows = Math.max(rows, this.lastRowOf(sh, c));
    } else {
      sh.unsure.forEach((c) => this.lastRowOf(sh, c));
      last.forEach((value, c) => {
        if (c >= left && c <= right && value > rows) rows = value;
      });
    }
    return rows;
  }

  /** How far each column of a sheet holds something (SheetState.lastRows), built in one pass over its cells. */
  private lastRows(sh: SheetState): Map<number, number> {
    if (sh.lastRows) return sh.lastRows;
    const last = new Map<number, number>();
    let looked = 0;
    sh.model.rows.forEach((row, index) => {
      if (!row || !row.cells) return;
      row.cells.forEach((cell, c) => {
        looked++;
        if (CalcEngine.holds(cell) && (last.get(c) || 0) < index + 1) last.set(c, index + 1);
      });
    });
    if (this.calculating) this.spend(looked >> 4);
    sh.lastRows = last;
    sh.unsure.clear();
    return last;
  }

  /** One past the last row of a column that holds something; a column whose last cell was emptied is looked at again, upwards. */
  private lastRowOf(sh: SheetState, col: number): number {
    const last = this.lastRows(sh);
    if (!sh.unsure.has(col)) return last.get(col) || 0;
    sh.unsure.delete(col);
    const keys = this.rowKeys(sh);
    const from = last.get(col) || 0;
    let found = 0;
    let looked = 0;
    for (let k = lowerBound(keys, from) - 1; k >= 0; k--) {
      looked++;
      const row = sh.model.rows.get(keys[k]);
      if (row && row.cells && CalcEngine.holds(row.cells[col])) {
        found = keys[k] + 1;
        break;
      }
    }
    if (this.calculating) this.spend(looked >> 4);
    if (found > 0) last.set(col, found);
    else last.delete(col);
    return found;
  }

  /**
   * Cells that were changed and put back without a recalculation (a value tried and taken back, see tryOps()): what
   * the engine keeps of where columns hold something follows them. Nothing is calculated.
   */
  cellsChanged(cells: CellAddress[]): void {
    for (const cell of cells) {
      const sh = this.sheets[cell.sheet];
      if (sh && isIndex(cell.row, MAX_ROWS) && isIndex(cell.col, MAX_COLS)) this.noteContent(sh, cell.row, cell.col);
    }
  }

  /** A cell the engine was told was changed (see recalculate()): how far its column holds something follows it. */
  private noteContent(sh: SheetState, row: number, col: number): void {
    sh.contentVersion++;
    const modelRow = sh.model.rows.get(row);
    if (sh.rowKeys) {
      const at = lowerBound(sh.rowKeys, row);
      if (!!modelRow !== (at < sh.rowKeys.length && sh.rowKeys[at] === row)) sh.rowKeys = undefined;
    }
    const last = sh.lastRows;
    if (!last) return;
    const before = last.get(col) || 0;
    if (CalcEngine.holds(modelRow ? modelRow.cells[col] : undefined)) {
      if (row + 1 > before) last.set(col, row + 1);
    } else if (row + 1 >= before && before > 0) sh.unsure.add(col);
  }

  /** One past the last column that holds something in the rows of `range` (all rows without one). */
  private contentCols(sh: SheetState, range?: CellRange): number {
    const top = range ? range.top : 0;
    const bottom = range ? range.bottom : MAX_ROWS - 1;
    const key = `c${top}:${bottom}`;
    if (!sh.extents || sh.extents.version !== sh.contentVersion) sh.extents = { version: sh.contentVersion, found: new Map() };
    const cache = sh.extents.found;
    const known = cache.get(key);
    if (known !== undefined) return known;
    let cols = 0;
    let looked = 0;
    const visit = (row: Row | undefined) => {
      if (!row || !row.cells) return;
      const cells = row.cells;
      for (let c = cells.length - 1; c >= cols; c--) {
        looked++;
        if (CalcEngine.holds(cells[c])) {
          cols = c + 1;
          break;
        }
      }
    };
    if (bottom - top < DIRECT_ROWS) for (let r = top; r <= bottom; r++) visit(sh.model.rows.get(r));
    else sh.model.rows.forEach((row, index) => {
      if (index >= top && index <= bottom) visit(row);
    });
    if (this.calculating) this.spend(looked >> 4);
    cache.set(key, cols);
    return cols;
  }

  /** The spills (written or not, of any array formula) that reach into `range` (see contentExtent()). */
  private reachOf(sh: SheetState, range: CellRange, visit: (spill: CellRange) => void): void {
    this.eachSpillMeeting(sh, range, (rec) => visit(rec.spill as CellRange));
    sh.foreign.forEach((spill) => {
      if (meets(spill, range)) visit(spill);
    });
  }

  private cellValue(sheet: number, row: number, col: number): Scalar {
    const sh = this.sheets[sheet];
    if (!sh) return null;
    if (++this.work >= this.nextCheck) this.spend(0);
    const modelRow = sh.model.rows.get(row);
    const cell = modelRow ? modelRow.cells[col] : undefined;
    if (!cell) {
      // A cell an array formula fills may not exist yet.
      if (sh.spillCells.size > 0 || sh.spillRanges.length > 0) {
        const owner = this.previewing ? undefined : this.spillOwner(sh, row, col);
        if (owner && owner.epoch === this.epoch) return this.spillValue(owner, row, col, undefined);
      }
      return null;
    }
    return this.valueOf(sh, cell, row, col);
  }

  private valueOf(sh: SheetState, cell: Cell, row: number, col: number): Scalar {
    if (this.previewing) return modelValue(cell, !!this.workbook.date1904);
    if (cell.formula !== undefined) {
      if (sh.rowEpoch[row] === this.epoch) {
        const rec = sh.recs.get(keyOf(row, col));
        if (rec !== undefined && rec.epoch === this.epoch && rec.formula === cell.formula) return this.need(rec);
      }
    } else {
      const owner = sh.spillCells.size > 0 || sh.spillRanges.length > 0 ? this.spillOwner(sh, row, col) : undefined;
      // What was typed into the cells of a dynamic array formula's spill is not its result: it blocks it.
      if (owner && owner.epoch === this.epoch && (!owner.dynamic || cell.spilled === true || isVacant(cell))) return this.spillValue(owner, row, col, cell);
      // A cell a spill left holds nothing any more, though it is emptied only when results are written; so does one no spill
      // fills at all (put back by an undo after the spill went), unless an array formula outside the graph fills it.
      if (cell.spilled === true && (!owner || owner.dynamic) && (owner ? this.vacated.length > 0 && this.isVacated(sh.index, row, col) : !this.isForeign(sh, row, col))) return null;
    }
    return modelValue(cell, !!this.workbook.date1904);
  }

  private isVacated(sheet: number, row: number, col: number): boolean {
    for (const area of this.vacated) {
      const range = area.range;
      if (area.sheet === sheet && row >= range.top && row <= range.bottom && col >= range.left && col <= range.right) return true;
    }
    return false;
  }

  private forEachCell(sheet: number, range: CellRange, visit: (value: Scalar, row: number, col: number) => void): void {
    const sh = this.sheets[sheet];
    if (!sh || !range) return;
    const top = Math.max(range.top, 0);
    const bottom = Math.min(range.bottom, MAX_ROWS - 1);
    const left = Math.max(range.left, 0);
    const right = Math.min(range.right, MAX_COLS - 1);
    if (!(top <= bottom && left <= right)) return;
    this.eachExisting(sh, top, left, bottom, right, visit);
    this.visitMissing(sh, top, left, bottom, right, visit);
  }

  private eachExisting(sh: SheetState, top: number, left: number, bottom: number, right: number, visit: (value: Scalar, row: number, col: number) => void): void {
    const rows = sh.model.rows;
    if (bottom - top < DIRECT_ROWS) {
      for (let r = top; r <= bottom; r++) {
        const row = rows.get(r);
        if (row) this.visitRow(sh, row, r, left, right, visit);
      }
      return;
    }
    const keys = this.rowKeys(sh);
    for (let k = lowerBound(keys, top); k < keys.length && keys[k] <= bottom; k++) {
      const row = rows.get(keys[k]);
      if (row) this.visitRow(sh, row, keys[k], left, right, visit);
    }
  }

  /**
   * Visits the cells of a range that an array formula fills and that do not exist yet: their values
   * are written only when the results are (see store()), and a range read before must meet them.
   */
  private visitMissing(sh: SheetState, top: number, left: number, bottom: number, right: number, visit: (value: Scalar, row: number, col: number) => void): void {
    if (this.previewing || sh.spills.size === 0) return;
    const rows = sh.model.rows;
    this.eachSpillMeeting(sh, { top, left, bottom, right }, (rec) => {
      const spill = rec.spill as CellRange;
      if (rec.epoch !== this.epoch) return;
      const r0 = Math.max(top, spill.top);
      const r1 = Math.min(bottom, spill.bottom);
      const c0 = Math.max(left, spill.left);
      const c1 = Math.min(right, spill.right);
      if (r0 > r1 || c0 > c1) return;
      this.spend((r1 - r0 + 1) * (c1 - c0 + 1));
      for (let r = r0; r <= r1; r++) {
        const row = rows.get(r);
        for (let c = c0; c <= c1; c++) {
          if (row && row.cells[c]) continue;
          // Visited once: the cell belongs to the first spill that covers it (spillOwner()).
          if (this.spillOwner(sh, r, c) !== rec) continue;
          const value = this.spillValue(rec, r, c, undefined);
          if (value !== null) visit(value, r, c);
        }
      }
    });
  }

  private visitRow(sh: SheetState, row: Row, index: number, left: number, right: number, visit: (value: Scalar, row: number, col: number) => void): void {
    const cells = row.cells;
    const last = Math.min(right, cells.length - 1);
    if (last >= left) this.spend(last - left + 1);
    for (let c = left; c <= last; c++) {
      const cell = cells[c];
      if (!cell) continue;
      const value = this.valueOf(sh, cell, index, c);
      if (value !== null) visit(value, index, c);
    }
  }

  private rowKeys(sh: SheetState): number[] {
    if (sh.rowKeys) return sh.rowKeys;
    const keys: number[] = [];
    let sorted = true;
    let previous = -1;
    sh.model.rows.forEach((_row, key) => {
      if (key < previous) sorted = false;
      previous = key;
      keys.push(key);
    });
    if (!sorted) keys.sort((a, b) => a - b);
    sh.rowKeys = keys;
    sh.rowKeysOf = sh.model.rows;
    return keys;
  }

  private spillOwner(sh: SheetState, row: number, col: number): FormulaRec | undefined {
    const listed = sh.spillCells.get(keyOf(row, col));
    if (listed) return listed;
    for (const rec of sh.spillRanges) {
      const range = rec.spill as CellRange;
      if (row >= range.top && row <= range.bottom && col >= range.left && col <= range.right) return rec;
    }
    return undefined;
  }

  private spillValue(rec: FormulaRec, row: number, col: number, cell: Cell | undefined): Scalar {
    this.need(rec);
    // The values of the file stay: this is one of them.
    if (rec.keep === KEEP_STORED) return cell ? modelValue(cell, !!this.workbook.date1904) : null;
    if (rec.circular || rec.looped) return 0;
    const range = rec.spill;
    // Calculated, the formula may no longer spill there: what it left holds nothing.
    if (!range || row < range.top || row > range.bottom || col < range.left || col > range.right) return cell && !(rec.dynamic && cell.spilled === true) ? modelValue(cell, !!this.workbook.date1904) : null;
    const value = arrayAt(rec.result || ZERO, row - range.top, col - range.left);
    return value === null ? 0 : value;
  }

  // -------------------------------------------------------------------------
  // Calculating
  //
  // A formula is calculated when something asks for its value, which is recursion: a chain
  // of 100,000 cells each reading the one before would overflow the stack. So the recursion
  // is cut short. When too many formulas are waiting on one another, the one that was about
  // to start is noted and everything on the stack is abandoned; the loop in calculate()
  // then starts again from that formula, with an empty stack, and works its way back up
  // the chain. A formula is calculated at most twice that way.
  //
  // Circular references. The formulas that read one another through the references they
  // name are known before anything is calculated (see components()): all of them get 0,
  // whether or not the part of a formula that closes the circle is ever calculated, and in
  // whatever order and however many slices the cells are calculated. A circle that only
  // exists through cells chosen while calculating (OFFSET, INDIRECT) shows when a formula is
  // asked for while it is being calculated: every formula waiting on the stack at that
  // moment is part of it, and none of them has a value yet that another formula could have
  // read. Formulas that only read a circle are calculated normally.

  /** The value of a dirty formula, calculating it first if needed. */
  private need(rec: FormulaRec): Scalar {
    if (rec.state === DONE) return rec.value;
    const chain = this.chain;
    if (rec.circular || rec.looped) {
      if (rec.state === ACTIVE) return 0;
      this.settle(rec);
      return rec.value;
    }
    if (rec.state === ACTIVE) {
      this.markLoop(rec);
      return 0;
    }
    if (chain.length - this.floor >= MAX_LIVE || evaluationDepth() > MAX_LIVE_DEPTH) {
      this.pending = rec;
      throw UNWIND;
    }
    const context = this.context;
    const sheet = context.sheet;
    const row = context.row;
    const col = context.col;
    rec.state = ACTIVE;
    rec.position = chain.length;
    chain.push(rec);
    const value = this.run(rec);
    this.finish(rec, value);
    context.sheet = sheet;
    context.row = row;
    context.col = col;
    return rec.value;
  }

  /** Calculates the dirty formulas from where the last call stopped, until all are done or the budget is used up. */
  private calculateAll(): void {
    const dirty = this.dirty;
    const epoch = this.epoch;
    this.calculating = true;
    try {
      const first = this.first;
      if (first) {
        this.first = undefined;
        if (first.state !== DONE && first.epoch === epoch && !first.gone) {
          this.sliceRoot = first;
          this.calculate(first);
          this.flush();
        }
      }
      for (; this.cursor < dirty.length; this.cursor++) {
        const rec = dirty[this.cursor];
        if (rec.state === DONE || rec.epoch !== epoch) continue;
        if (!this.sliceRoot) this.sliceRoot = rec;
        this.calculate(rec);
        // Written while the clock runs, so that writing is part of the budget. Nothing is being
        // calculated at this point, and what is written is final.
        this.flush();
      }
    } catch (error) {
      this.abandon();
    }
    this.calculating = false;
  }

  /** Out of budget: what is not done is left for the next call. */
  private abandon(): void {
    const chain = this.chain;
    const blamed = chain.length > 0 ? chain[chain.length - 1] : undefined;
    for (const rec of chain) {
      if (rec.state === DONE) continue;
      rec.state = FRESH;
      rec.result = undefined;
    }
    chain.length = 0;
    this.pending = undefined;
    if (!blamed || blamed.state === DONE) return;
    if (this.finished === 0 && blamed === this.sliceRoot) {
      // It had the hard budget to itself. It keeps its value until something it reads changes.
      blamed.costly = true;
      blamed.epoch = 0;
      blamed.live = false;
      this.costlyCells.set(`${blamed.sheet}:${blamed.row}:${blamed.col}`, { sheet: blamed.sheet, row: blamed.row, col: blamed.col });
      this.staleCells = undefined;
    } else this.first = blamed;
  }

  private calculate(root: FormulaRec): void {
    if (root.state === DONE) return;
    const chain = this.chain;
    // Left over from a calculation that was cut off by an exception.
    for (const rec of chain) if (rec.state === ACTIVE) rec.state = FRESH;
    chain.length = 0;
    if (root.circular || root.looped) {
      this.settle(root);
      return;
    }
    root.state = ACTIVE;
    root.position = 0;
    chain.push(root);
    while (chain.length > 0) {
      const top = chain[chain.length - 1];
      this.floor = chain.length;
      let value: Scalar;
      try {
        value = this.run(top);
      } catch (error) {
        if (error === UNWIND && this.pending) {
          // The formulas that were waiting stay on the chain; the one they wait for goes on top.
          const next = this.pending;
          this.pending = undefined;
          next.state = ACTIVE;
          next.position = chain.length;
          chain.push(next);
          continue;
        }
        if (error === BUDGET) throw error;
        value = ERROR_VALUE;
      }
      this.finish(top, value);
    }
  }

  /** Evaluates one formula. Only the signals to start over or to give up get out. */
  private run(rec: FormulaRec): Scalar {
    const context = this.context;
    context.sheet = rec.sheet;
    context.row = rec.row;
    context.col = rec.col;
    this.spend(1);
    const modelRow = this.sheets[rec.sheet].model.rows.get(rec.row);
    const cell = modelRow ? modelRow.cells[rec.col] : undefined;
    if (!cell || cell.formula !== rec.formula) {
      if (!rec.gone) {
        rec.gone = true;
        this.outdated.push(rec);
      }
      return cell ? modelValue(cell, !!this.workbook.date1904) : null;
    }
    this.evalOptions.rowShift = rec.dRow;
    this.evalOptions.colShift = rec.dCol;
    this.evalOptions.arrayFormula = rec.arrayMode;
    if (rec.dynamic) return this.runDynamic(rec);
    // A formula calculated while a dynamic array formula that needs it is: it is no dynamic one itself.
    this.evalOptions.dynamic = false;
    try {
      const value = evaluate(rec.node, context, this.functions, this.evalOptions);
      if (!rec.isArray) return rec.arrayMode ? this.firstOf(value) : toCellValue(value, context);
      let result: ArrayValue;
      if (isRef(value)) {
        // Only as much of the reference as the formula's cells show is read: they are all the result is asked for.
        const area = rec.spill || { top: rec.row, left: rec.col, bottom: rec.row, right: rec.col };
        const range = value.range;
        const bottom = range.bottom >= range.top ? Math.min(range.bottom, range.top + area.bottom - area.top) : range.bottom;
        const right = range.right >= range.left ? Math.min(range.right, range.left + area.right - area.left) : range.right;
        result = refToArray({ ...value, range: { top: range.top, left: range.left, bottom, right } }, context);
      } else result = isArray(value) ? value : { rows: 1, cols: 1, data: [value] };
      rec.result = result;
      return result.data.length > 0 ? result.data[0] : ERROR_VALUE;
    } catch (error) {
      if (error instanceof EvaluationAbort) throw error;
      return ERROR_VALUE;
    }
  }

  /**
   * Evaluates a dynamic array formula and works out where its result goes (see SpillPlan):
   * the cells below and to the right of its own, as many as the result has, when nothing is
   * in them. Nothing is written or registered here: conclude() does that, once the value stands.
   */
  private runDynamic(rec: FormulaRec): Scalar {
    const context = this.context;
    this.evalOptions.dynamic = true;
    let value: Value;
    try {
      value = evaluate(rec.node, context, this.functions, this.evalOptions);
    } catch (error) {
      if (error instanceof EvaluationAbort) throw error;
      value = ERROR_VALUE;
    } finally {
      this.evalOptions.dynamic = false;
    }
    const own: CellRange = { top: rec.row, left: rec.col, bottom: rec.row, right: rec.col };
    let rows = 1;
    let cols = 1;
    let result: ArrayValue | undefined;
    if (isRef(value)) {
      // A whole column or row spills as far as its sheet holds something (see readRange()).
      const range = readRange(value.range, value.sheet, context, Infinity);
      rows = range.bottom - range.top + 1;
      cols = range.right - range.left + 1;
      if (rows === 1 && cols === 1) value = this.cellValue(value.sheet, range.top, range.left);
      // The whole reference spills, empty cells included (as 0).
      else if (rows * cols <= this.spillLimit) result = refToArray({ ...value, range }, context);
    } else if (isArray(value)) {
      rows = value.rows;
      cols = value.cols;
      if (rows === 1 && cols === 1) value = value.data[0];
      else result = value;
    }
    if (rows === 1 && cols === 1) {
      rec.next = { want: own, blocked: false };
      return value as Scalar;
    }
    // Cut to the sheet, so that the formula watches the cells that are there.
    const want: CellRange = { top: rec.row, left: rec.col, bottom: Math.min(rec.row + rows - 1, MAX_ROWS - 1), right: Math.min(rec.col + cols - 1, MAX_COLS - 1) };
    // UNSURE: Excel's error for a result too large to spill was not checked; #SPILL! is what it gives at the edge of the sheet.
    const fits = rec.row + rows <= MAX_ROWS && rec.col + cols <= MAX_COLS && rows * cols <= this.spillLimit && !!result;
    // Taken up too often: the spills around it come and go with one another (see resolve()).
    const tooOften = rec.reopenEpoch === this.epoch && rec.reopens >= MAX_REOPENS;
    if (tooOften) this.tooOften.add(rec);
    if (!fits || tooOften || this.blocked(rec, want)) {
      rec.next = { want, blocked: true };
      return calcError('#SPILL!');
    }
    // It wants the cells; whether it gets them within the budget of all spills is worked out from what all of them
    // want, in their order on the sheets (see allocate()).
    if (this.budgetRefused.has(rec) && !this.fileSpillStays(rec, want)) {
      rec.next = { want, blocked: true, budget: true };
      return calcError('#SPILL!');
    }
    rec.next = { want, spill: want, blocked: false, result, budget: true };
    return (result as ArrayValue).data.length > 0 ? (result as ArrayValue).data[0] : ERROR_VALUE;
  }

  /**
   * The spills, once everything else is calculated (see recalculate()): circles that run through them, spills that
   * come and go with one another (a formula taken up too often), and the budget of all spills. It runs to its end
   * whatever the budget of the call: what it gives must not depend on the slices a recalculation is cut into.
   */
  private settleSpills(): void {
    const saved = [this.workLimit, this.timeLimit, this.softWork, this.softTime, this.hardWork, this.hardTime, this.nextCheck];
    this.workLimit = this.timeLimit = this.softWork = this.softTime = this.hardWork = this.hardTime = this.nextCheck = Infinity;
    try {
      for (let round = 0; round < MAX_CIRCLE_ROUNDS; round++) {
        if (this.seeds.length > 0 || this.search) this.components();
        const flipped = this.takeFlipped();
        const suspects = new Set<FormulaRec>();
        this.tooOften.forEach((rec) => {
          if (!rec.gone) suspects.add(rec);
        });
        this.tooOften.clear();
        // A dynamic array formula met in a circle while calculating (through a spill it or another one made): which
        // of them met it depends on the order they came in; what they give is worked out the same way as above.
        for (const rec of this.loops) if (rec.looped && rec.dynamic && !rec.gone && !(rec.forced && rec.forcedEpoch === this.epoch)) suspects.add(rec);
        for (const rec of flipped) {
          if (rec.gone) continue;
          if (rec.dynamic) suspects.add(rec);
          // The other formulas of its circle that spill.
          const members = rec.comp ? this.groups.get(rec.comp) : undefined;
          if (members) for (const member of members) if (member.dynamic && !member.gone) suspects.add(member);
        }
        if (suspects.size > 0) this.resolve(suspects);
        else if (flipped.length > 0) {
          this.retake(flipped);
          this.calculateAll();
          this.flush();
        }
        // The budget of all spills, shared out by what they want now; its changes may make or break circles again.
        const allocated = this.allocate();
        if (suspects.size === 0 && flipped.length === 0 && !allocated && this.tooOften.size === 0 && this.seeds.length === 0) break;
      }
    } finally {
      [this.workLimit, this.timeLimit, this.softWork, this.softTime, this.hardWork, this.hardTime, this.nextCheck] = saved;
    }
  }

  /**
   * Dynamic array formulas whose spills make a circle (a formula reads the cells another one spills into, which reads
   * it), or come and go with one another: what they give is worked out from no spill at all, the same way whatever
   * happened before. In each round every one of them is calculated with the spills of the round before, and their
   * spills change together at its end; when a round changes none, that is what they give. When the rounds do not
   * settle, or the spills they settle on still make a circle, all of them are circular: 0 in their own cells, as
   * Excel leaves the formulas of a circular reference, and they spill nowhere. UNSURE: Excel says a circular reference
   * there; its values were not compared.
   */
  private resolve(suspects: Set<FormulaRec>): void {
    const members = Array.from(suspects);
    // What reads the cells of a spill that changes is calculated again with it.
    const touched: FormulaRec[] = [];
    const leave = (rec: FormulaRec) => {
      const sh = this.sheets[rec.sheet];
      if (!rec.spill) return;
      for (const reader of this.readersOf(rec.sheet, rec.spill)) touched.push(reader);
      this.unregisterSpill(sh, rec);
      sh.columns = undefined;
    };
    // From no spill at all, and without what circles were found from the spills of before, nor the share of the
    // budget they had: the budget is shared out again afterwards (see allocate()).
    for (const rec of members) {
      leave(rec);
      this.clearCircle(rec);
      this.budgetRefused.delete(rec);
    }
    let settled = false;
    let moving: FormulaRec[] = [];
    const circular = new Set<FormulaRec>();
    const force = (rec: FormulaRec) => {
      if (circular.has(rec)) return;
      circular.add(rec);
      leave(rec);
      rec.circular = true;
      rec.forced = true;
      rec.forcedEpoch = this.epoch;
      this.sheets[rec.sheet].forcedRecs.add(rec);
      touched.push(rec);
    };
    for (let round = 0; round < MAX_SETTLE_ROUNDS && !settled; round++) {
      this.deferred = new Set();
      // What was found circular while calculating with the spills of a round holds for that round only.
      for (const rec of this.loops) rec.looped = false;
      this.loops = [];
      this.retake(members.concat(touched));
      touched.length = 0;
      this.calculateAll();
      const moved = Array.from(this.deferred);
      this.deferred = undefined;
      settled = true;
      moving = [];
      for (const rec of moved) {
        if (rec.gone || !rec.next || circular.has(rec)) continue;
        const sh = this.sheets[rec.sheet];
        const after = rec.next.spill;
        if (sameRange(rec.spill, after)) continue;
        leave(rec);
        if (after) {
          this.registerSpill(sh, rec, after);
          for (const reader of this.readersOf(rec.sheet, after)) touched.push(reader);
        }
        settled = false;
        moving.push(rec);
        if (!suspects.has(rec)) {
          suspects.add(rec);
          members.push(rec);
        }
      }
      if (settled) break;
      // The circles the spills of this round make: their formulas are circular from now on, 0 and spilling nowhere,
      // and the others go on without them. Circular is circular for good in these rounds, so they end.
      for (const rec of members) this.seeds.push(rec);
      this.components();
      for (const rec of this.takeFlipped()) {
        if (rec.circular && suspects.has(rec)) force(rec);
        // Out of a circle (the spill it ran through is gone): calculated again, not left with its circular 0.
        else if (!rec.circular && !rec.gone) touched.push(rec);
      }
      for (const rec of members) if (rec.circular && !rec.gone) force(rec);
    }
    // Spills that still come and go after the last round: their formulas are circular as well.
    if (!settled) for (const rec of moving) force(rec);
    // Circles the spills they settled on still make.
    for (const rec of members) this.seeds.push(rec);
    this.components();
    // A formula that is out of a circle now (the spills of the circular ones are gone) is calculated with them gone.
    for (const rec of this.takeFlipped()) if (!rec.circular && !rec.gone) touched.push(rec);
    for (const rec of members) if (rec.circular && !rec.gone) force(rec);
    // They are settled again together when any of them changes or goes (see clearCircle()).
    const group = Array.from(circular);
    for (const rec of group) rec.forcedGroup = group;
    // A forced circle is not looked for again in this recalculation: no seeds for it.
    this.retake(touched.slice());
    touched.length = 0;
    this.calculateAll();
    this.flush();
  }

  /**
   * A circle that runs through spills, as the file was saved with it (see resolve()): its dynamic array formulas are 0
   * in their own cells and fill nothing, though they would spill if calculated. Such a formula, with the cells it would
   * spill into, is put back into the graph as circular; if a circle runs through them, it is the one the engine that
   * saved the file found, and it is settled again when something in it changes, as it would be there. The values of
   * the file stay: nothing is calculated for it. A formula that is in no circle is what it was (a value of the file).
   */
  private rebuildForced(): void {
    const list = this.reguarded;
    this.reguarded = [];
    const found: FormulaRec[] = [];
    for (const rec of list) {
      if (rec.gone || !rec.dynamic || rec.spill || rec.forced || rec.live || !rec.guard || spillSize(rec.guard.range) <= 1) continue;
      if (!rec.next || !rec.next.spill || this.storedOf(rec) !== 0) continue;
      found.push(rec);
    }
    // A file with more of them than any workbook made by hand: its values stay, and no circle is found again.
    if (found.length > MAX_REBUILT) return;
    const candidates: FormulaRec[] = [];
    for (const rec of found) {
      rec.forced = true;
      rec.forcedEpoch = 0;
      rec.circular = true;
      this.sheets[rec.sheet].forcedRecs.add(rec);
      candidates.push(rec);
    }
    if (candidates.length === 0) return;
    const saved = [this.workLimit, this.timeLimit, this.softWork, this.softTime, this.hardWork, this.hardTime, this.nextCheck];
    this.workLimit = this.timeLimit = this.softWork = this.softTime = this.hardWork = this.hardTime = this.nextCheck = Infinity;
    this.quiet = true;
    try {
      for (const rec of candidates) this.seeds.push(rec);
      this.components();
    } finally {
      this.quiet = false;
      [this.workLimit, this.timeLimit, this.softWork, this.softTime, this.hardWork, this.hardTime, this.nextCheck] = saved;
    }
    for (const rec of candidates) rec.forcedGroup = candidates;
  }

  private unforce(rec: FormulaRec): void {
    rec.forced = false;
    const sh = this.sheets[rec.sheet];
    sh.forcedRecs.delete(rec);
  }

  /** Takes a formula out of the circle it was found in (see resolve()). */
  private clearCircle(rec: FormulaRec): void {
    const group = rec.forcedGroup;
    if (rec.forced) this.unforce(rec);
    rec.forcedGroup = undefined;
    if (group) {
      for (const member of group) {
        if (member === rec || !member.forced || member.gone) continue;
        this.unforce(member);
        member.forcedGroup = undefined;
        member.circular = member.selfRef;
        this.tooOften.add(member);
        this.seeds.push(member);
      }
    }
    if (rec.comp) {
      const members = this.groups.get(rec.comp);
      this.groups.delete(rec.comp);
      if (members) for (const member of members) {
        member.comp = 0;
        member.circular = member.selfRef;
      }
    }
    rec.comp = 0;
    rec.circular = rec.selfRef;
  }

  /**
   * Whether a dynamic array formula keeps the spill the FILE gave it whatever the budget: the very cells of its `ref`
   * as the file had it, all of them there (a damaged file may name a million cells and hold none). They are in the
   * workbook already. A `ref` the engine wrote does not count.
   */
  private fileSpillStays(rec: FormulaRec, want: CellRange): boolean {
    // The cell as the file has it (not one an engine wrote since), with that very `ref`.
    const modelRow = this.sheets[rec.sheet].model.rows.get(rec.row);
    const cell = modelRow ? modelRow.cells[rec.col] : undefined;
    if (!cell || cell.formula !== rec.formula || this.own.has(cell) || !sameRange(spillAreaOf(cell.formula, rec.row, rec.col), want)) return false;
    let present = 0;
    this.eachCell(this.sheets[rec.sheet], want, (cell) => {
      if (cell.spilled === true || cell.formula) present++;
    });
    this.spend(present);
    return present >= spillSize(want);
  }

  /**
   * The budget of all spills (MAX_WORKBOOK_SPILL_CELLS), shared out by place: the dynamic array formulas that want
   * more than one cell, in the order of sheets, rows and columns, get what they want while it lasts; the first one
   * past it and the ones after it that do not fit are #SPILL!. What a formula wants is what it gives, whatever the
   * budget; formulas of one value count for nothing; a spill the file gave stays (see fileSpillStays()). Who spills
   * so depends on what the formulas give, not on the order they were calculated in. Those whose lot changed are
   * calculated again, a few times at most (a spill a formula reads can change what an earlier one wants).
   */
  private allocate(): boolean {
    const before = this.budgetRefused;
    const lot = () => {
      if (this.budgetRefused.size !== before.size) return true;
      for (const rec of this.budgetRefused) if (!before.has(rec)) return true;
      return false;
    };
    // Shared out from no refusal, once a recalculation: what was refused before does not decide who is refused now
    // (two lots can both hold when what an earlier formula wants depends on whether a later one spills). It costs
    // the refused spills being worked out (and written) once more in each recalculation that has any.
    if (before.size > 0 && this.allocatedEpoch !== this.epoch) {
      this.budgetRefused = new Set();
      this.retake(Array.from(before).filter((rec) => !rec.gone));
      this.calculateAll();
      this.flush();
    }
    this.allocatedEpoch = this.epoch;
    for (let round = 0; round < MAX_ALLOCATE_ROUNDS; round++) {
      const wanting: { rec: FormulaRec; want: CellRange; size: number }[] = [];
      let total = 0;
      this.wanting.forEach((rec) => {
        if (rec.gone) {
          this.wanting.delete(rec);
          return;
        }
        // A formula found circular through its spill (see resolve(), rebuildForced()) fills nothing until it is calculated
        // anew, though the plan it was put back with says where it would spill: it takes nothing of the budget either.
        if (!rec.dynamic || rec.circular || rec.forced) return;
        const plan = rec.next;
        const want = plan && plan.budget ? plan.want : rec.spill;
        if (!want) return;
        const size = spillSize(want);
        if (size <= 1) return;
        wanting.push({ rec, want, size });
        total += size;
      });
      const refused = new Set<FormulaRec>();
      if (total > this.bookLimit) {
        wanting.sort((a, b) => a.rec.sheet - b.rec.sheet || a.rec.row - b.rec.row || a.rec.col - b.rec.col);
        let used = 0;
        for (const item of wanting) {
          if (used + item.size <= this.bookLimit || this.fileSpillStays(item.rec, item.want)) used += item.size;
          else refused.add(item.rec);
        }
      }
      const changed: FormulaRec[] = [];
      this.budgetRefused.forEach((rec) => {
        if (!refused.has(rec)) changed.push(rec);
      });
      refused.forEach((rec) => {
        if (!this.budgetRefused.has(rec)) changed.push(rec);
      });
      this.budgetRefused = refused;
      if (changed.length === 0) return lot();
      this.retake(changed);
      this.calculateAll();
      this.flush();
    }
    return lot();
  }

  /**
   * Whether the spill of `other` gives way to the one `rec` wants: two dynamic array formulas that would spill into
   * the same cells are settled by place, the first on the sheet (by row, then column) spills and the other is #SPILL!,
   * whichever was calculated first. UNSURE: Excel gives #SPILL! to the one it calculates last.
   */
  private yields(other: FormulaRec, rec: FormulaRec): boolean {
    return other.dynamic && rec.dynamic && (other.row > rec.row || (other.row === rec.row && other.col > rec.col));
  }

  /** Whether anything is in the cells a dynamic array formula would spill into: a value, a formula, a merged cell, another spill. */
  private blocked(rec: FormulaRec, area: CellRange): boolean {
    const sh = this.sheets[rec.sheet];
    const model = sh.model;
    for (const merge of model.merges || []) {
      if (merge && merge.top <= area.bottom && merge.bottom >= area.top && merge.left <= area.right && merge.right >= area.left) return true;
    }
    // The cells of an array formula that is not calculated here, however empty they are.
    if (sh.foreign.size > 0) {
      for (const range of sh.foreign.values()) {
        if (range.top <= area.bottom && range.bottom >= area.top && range.left <= area.right && range.right >= area.left) return true;
      }
    }
    const check = (row: Row | undefined, r: number): boolean => {
      if (!row || !row.cells) return false;
      const cells = row.cells;
      const last = Math.min(area.right, cells.length - 1);
      if (last >= area.left) this.spend(last - area.left + 1);
      for (let c = area.left; c <= last; c++) {
        const cell = cells[c];
        if (!cell || (r === rec.row && c === rec.col) || cell.spilled === true) continue;
        if (!isVacant(cell)) return true;
      }
      return false;
    };
    if (area.bottom - area.top < DIRECT_ROWS) {
      for (let r = area.top; r <= area.bottom; r++) if (check(model.rows.get(r), r)) return true;
    } else {
      const keys = this.rowKeys(sh);
      for (let k = lowerBound(keys, area.top); k < keys.length && keys[k] <= area.bottom; k++) if (check(model.rows.get(keys[k]), keys[k])) return true;
    }
    // The cells of another spill, written or not yet.
    let hit = false;
    this.eachSpillMeeting(sh, area, (other) => {
      if (other !== rec && !this.yields(other, rec)) hit = true;
    });
    if (hit) return true;
    return false;
  }

  /**
   * The dynamic array formula whose result fills a cell (its own cell or one it spills into),
   * as the last recalculation left it: where the formula is and the cells it fills. Undefined
   * for any other cell, and while the graph is not built.
   */
  spillAt(cell: CellAddress): { sheet: number; row: number; col: number; range: CellRange } | undefined {
    if (!cell || !this.built) return undefined;
    const sh = this.sheets[cell.sheet];
    if (!sh || !isIndex(cell.row, MAX_ROWS) || !isIndex(cell.col, MAX_COLS)) return undefined;
    const own = sh.recs.get(keyOf(cell.row, cell.col));
    const rec = own && own.dynamic && own.spill ? own : sh.spillCells.size > 0 || sh.spillRanges.length > 0 ? this.spillOwner(sh, cell.row, cell.col) : undefined;
    if (!rec || !rec.dynamic || !rec.spill || rec.gone) return undefined;
    return { sheet: rec.sheet, row: rec.row, col: rec.col, range: { ...rec.spill } };
  }

  /** The cells the dynamic array formula of a cell spills into (CalcContext.spillRange). */
  private spillRangeOf(sheet: number, row: number, col: number): CellRange | CalcError | undefined {
    const sh = this.sheets[sheet];
    const rec = sh ? sh.recs.get(keyOf(row, col)) : undefined;
    if (!rec || !rec.dynamic || rec.gone) {
      // A dynamic array formula the engine does not calculate (an unknown function, a reference into another
      // workbook) is one all the same: it fills what the file says, or its own cell. Anything else has no spill.
      const modelRow = sh ? sh.model.rows.get(row) : undefined;
      const cell = modelRow ? modelRow.cells[col] : undefined;
      if (!cell || !cell.formula || cell.formula.dynamic !== true) return undefined;
      if (cell.kind === 'error' && cell.value === '#SPILL!') return undefined;
      return spillAreaOf(cell.formula, row, col) || { top: row, left: col, bottom: row, right: col };
    }
    if (!this.previewing && rec.epoch === this.epoch) this.need(rec);
    // UNSURE: a reference to a formula that does not spill (blocked, or a single value) is #REF! here; not checked against Excel.
    if (rec.keep !== KEEP_STORED && rec.state === DONE && rec.epoch === this.epoch && rec.next && rec.next.blocked) return undefined;
    if (rec.spill) return rec.spill;
    const value = rec.epoch === this.epoch && rec.state === DONE ? rec.value : this.cellValue(sheet, row, col);
    return isError(value) && value.error === '#SPILL!' ? undefined : { top: row, left: col, bottom: row, right: col };
  }

  /** What the cell of a dynamic array formula shows: the first element of its result. */
  private firstOf(value: Value): Scalar {
    if (isRef(value)) return this.cellValue(value.sheet, value.range.top, value.range.left);
    if (isArray(value)) return value.data.length > 0 ? value.data[0] : ERROR_VALUE;
    return value;
  }

  private finish(rec: FormulaRec, value: Scalar): void {
    const chain = this.chain;
    // Normally `rec` is the last one; anything above it was cut off by an exception.
    for (let k = chain.length - 1; k > rec.position; k--) if (chain[k].state === ACTIVE) chain[k].state = FRESH;
    chain.length = rec.position;
    if (rec.circular || rec.looped) {
      value = 0;
      rec.result = undefined;
      this.reportCircular(rec);
    }
    this.conclude(rec, value);
  }

  /** A formula known to be circular before it is calculated: it is 0. */
  private settle(rec: FormulaRec): void {
    rec.result = undefined;
    this.reportCircular(rec);
    this.conclude(rec, 0);
  }

  private reportCircular(rec: FormulaRec): void {
    const key = `${rec.sheet}:${rec.row}:${rec.col}`;
    if (this.circularKeys.has(key)) return;
    this.circularKeys.add(key);
    this.circular.push({ sheet: rec.sheet, row: rec.row, col: rec.col });
  }

  /**
   * Decides what a calculated formula is worth. When nothing it reads has changed and the
   * file has a value for it, that value stands: the same one is not written again, a
   * different one is kept and the cell reported.
   */
  private conclude(rec: FormulaRec, value: Scalar): void {
    // An empty result shows as 0, as a reference to an empty cell does.
    if (value === null || value === undefined) value = 0;
    else if (typeof value === 'number' && !isFinite(value)) value = ERROR_NUM;
    let keep = KEEP_NONE;
    let plan: SpillPlan | undefined;
    if (rec.dynamic && !rec.gone) {
      const own = { top: rec.row, left: rec.col, bottom: rec.row, right: rec.col };
      // A circular one is 0 in its own cell and spills nowhere, as the formulas of a circular reference in Excel; "#"
      // names that cell. Which formulas are circular when the circle runs through spills is settled by resolve().
      // It goes on watching where it would spill (its guard): what comes into the way there changes its circle.
      if (rec.circular || rec.looped) plan = { want: rec.forced && rec.guard ? rec.guard.range : own, blocked: false };
      else plan = rec.next || { want: own, blocked: false };
      rec.next = plan;
      rec.result = plan.result;
    }
    if (!rec.live) {
      const stored = this.storedOf(rec);
      if (stored !== undefined) {
        const sameSpill = !plan || sameRange(plan.spill, rec.spill);
        keep = sameValue(stored, value) && sameSpill && !(rec.spill && this.spillDiffers(rec)) ? KEEP_SAME : KEEP_STORED;
        value = stored;
      }
    }
    // Where it would spill is watched even while the value of the file stands: what comes out of the way changes it.
    if (plan) {
      this.setGuard(rec, plan.want);
      if ((plan.budget && spillSize(plan.want) > 1) || rec.spill) this.wanting.add(rec);
      else this.wanting.delete(rec);
    }
    if (plan && keep !== KEEP_STORED) this.applySpill(rec, plan);
    if (keep === KEEP_STORED) {
      if (!rec.kept) {
        rec.kept = true;
        this.mismatchedCells.set(`${rec.sheet}:${rec.row}:${rec.col}`, { sheet: rec.sheet, row: rec.row, col: rec.col });
      }
    } else if (rec.kept) {
      rec.kept = false;
      this.mismatchedCells.delete(`${rec.sheet}:${rec.row}:${rec.col}`);
    }
    rec.keep = keep;
    rec.value = value;
    rec.state = DONE;
    this.done.push(rec);
    if (rec.stale) {
      rec.stale = false;
      const walk = this.walk;
      this.walk = [];
      this.takeUp([rec], undefined, walk);
    }
    if (++this.finished === 1 && this.extended) {
      // The formula the budget was extended for is done: the next look at the budget ends the call.
      this.workLimit = this.softWork;
      this.timeLimit = this.softTime;
      this.nextCheck = 0;
    }
  }

  /**
   * A dynamic array formula found where its result goes: the formula watches those cells from
   * now on, and when it spills into other cells than before, the formulas that read the cells
   * it reaches or leaves are calculated (again, when they were done already).
   */
  private applySpill(rec: FormulaRec, plan: SpillPlan): void {
    const sh = this.sheets[rec.sheet];
    this.setGuard(rec, plan.want);
    const before = rec.spill;
    const after = plan.spill;
    if (sameRange(before, after)) return;
    // While spills are being settled, all of them change together, at the end of a round (see resolve()).
    if (this.deferred) {
      this.deferred.add(rec);
      return;
    }
    this.unregisterSpill(sh, rec);
    if (after) this.registerSpill(sh, rec, after);
    // The circles the spill makes or breaks are looked for once this round of calculating is over (see recalculate()).
    sh.columns = undefined;
    this.seeds.push(rec);
    const top = Math.min(before ? before.top : rec.row, after ? after.top : rec.row);
    const left = Math.min(before ? before.left : rec.col, after ? after.left : rec.col);
    const bottom = Math.max(before ? before.bottom : rec.row, after ? after.bottom : rec.row);
    const right = Math.max(before ? before.right : rec.col, after ? after.right : rec.col);
    this.reopen(rec, { top, left, bottom, right });
  }

  /**
   * The formulas that read a range whose values the spill of `from` changed: those not dirty
   * yet are marked, those already calculated in this recalculation are taken up again, and the
   * same for whatever reads them. A formula is taken up again a few times at most (see MAX_REOPENS).
   */
  private reopen(from: FormulaRec, area: CellRange): void {
    const sh = this.sheets[from.sheet];
    const walk = this.walk;
    const found = new Set<FormulaRec>();
    this.walk = [];
    this.gathering = found;
    this.collect(sh, area.top, area.left, area.bottom, area.right, true);
    const pending = Array.from(found);
    found.clear();
    this.gathering = undefined;
    this.walk = [];
    this.takeUp(pending, from, walk);
  }

  /** The formulas that read cells of `area`, not counted as taken up again (see resolve()). */
  private readersOf(sheet: number, area: CellRange): FormulaRec[] {
    const sh = this.sheets[sheet];
    const walk = this.walk;
    const found = new Set<FormulaRec>();
    this.walk = [];
    this.gathering = found;
    this.collect(sh, area.top, area.left, area.bottom, area.right, true);
    this.gathering = undefined;
    this.walk = walk;
    return Array.from(found);
  }

  /** The formulas whose circularity changed: calculated again in this recalculation (see reopen()). */
  private retake(list: FormulaRec[]): void {
    if (list.length === 0) return;
    const walk = this.walk;
    this.walk = [];
    this.takeUp(list.slice(), undefined, walk, false);
  }

  private takeFlipped(): FormulaRec[] {
    const list = this.flipped;
    this.flipped = [];
    return list;
  }

  /** Calculates formulas again, and what reads them (see reopen()); `walk` is put back when done. */
  private takeUp(pending: FormulaRec[], from: FormulaRec | undefined, walk: FormulaRec[], counted = true): void {
    const found = new Set<FormulaRec>();
    const epoch = this.epoch;
    // A formula that is to be calculated (again): what already read its value of before is taken up as well, and
    // each formula is looked at once.
    const seen = new Set<FormulaRec>();
    const readers = (rec: FormulaRec) => {
      this.gathering = found;
      this.dependents(rec, true);
      this.gathering = undefined;
      found.forEach((other) => pending.push(other));
      found.clear();
    };
    // What functions kept for this recalculation (an index of a range) may hold values of before.
    if (pending.length > 0) this.memos.clear();
    for (let k = 0; k < pending.length; k++) {
      const rec = pending[k];
      if (rec === from || rec.gone || seen.has(rec)) continue;
      seen.add(rec);
      if (rec.epoch !== epoch) {
        this.mark(rec, true);
        // Marked now, it may have been read already in this recalculation, its value of before.
        readers(rec);
        continue;
      }
      if (rec.state !== DONE) {
        rec.live = true;
        // Being calculated, it may have read the cells already: it is calculated again once it is done.
        if (rec.state === ACTIVE) rec.stale = true;
        readers(rec);
        continue;
      }
      if (rec.reopenEpoch !== epoch) {
        rec.reopenEpoch = epoch;
        rec.reopens = 0;
      }
      // A dynamic array formula is taken up again a few times at most, and is #SPILL! then (see runDynamic()): its
      // spill stops changing, and so does what reads it. Any other formula is taken up again whenever a spill it reads
      // changed, so that no participant is left with a value of before, whatever order formulas came in.
      if (counted) {
        if (rec.dynamic && rec.reopens > MAX_REOPENS) {
          // It keeps a value of before for now: tried again once this round is over (see recalculate()).
          this.tooOften.add(rec);
          continue;
        }
        rec.reopens++;
      } else {
        // Calculated again for a reason of its own (a circle found, see recalculate()): it may be taken up as often again.
        rec.reopenEpoch = epoch;
        rec.reopens = 0;
      }
      rec.state = FRESH;
      rec.keep = KEEP_NONE;
      rec.live = true;
      rec.result = undefined;
      // Found circular when spills were settled in an earlier recalculation: what it reads changed, so its circle is
      // settled again (see touch()).
      if (rec.forced && rec.forcedEpoch !== epoch) {
        this.clearCircle(rec);
        this.seeds.push(rec);
        this.tooOften.add(rec);
      }
      this.dirty.push(rec);
      // What read its old value is taken up as well.
      readers(rec);
    }
    // Formulas newly marked: their dependents are marked in turn.
    const marked = this.walk;
    for (let k = 0; k < marked.length; k++) this.dependents(marked[k], marked[k].live);
    this.walk = walk;
  }

  /** The value the file has for a formula; undefined when the cell has none or the engine wrote it. */
  private storedOf(rec: FormulaRec): Scalar | undefined {
    const modelRow = this.sheets[rec.sheet].model.rows.get(rec.row);
    const cell = modelRow ? modelRow.cells[rec.col] : undefined;
    if (!cell || cell.formula !== rec.formula || this.own.has(cell)) return undefined;
    const value = modelValue(cell, !!this.workbook.date1904);
    return value === null ? undefined : value;
  }

  /** Whether the cell of a formula holds no value at all. */
  private unvalued(rec: FormulaRec): boolean {
    const modelRow = this.sheets[rec.sheet].model.rows.get(rec.row);
    const cell = modelRow ? modelRow.cells[rec.col] : undefined;
    return !cell || cell.formula !== rec.formula || !hasValue(cell);
  }

  /** Whether a cell an array formula fills holds a value of the file other than the one calculated. */
  private spillDiffers(rec: FormulaRec): boolean {
    const sh = this.sheets[rec.sheet];
    const range = rec.spill as CellRange;
    const result = rec.result || ZERO;
    const zero = rec.circular || rec.looped;
    const date1904 = !!this.workbook.date1904;
    for (let r = range.top; r <= range.bottom; r++) {
      const row = sh.model.rows.get(r);
      if (!row) continue;
      for (let c = range.left; c <= range.right; c++) {
        const cell = row.cells[c];
        if (!cell || cell.formula || this.own.has(cell) || (r === rec.row && c === rec.col)) continue;
        const stored = modelValue(cell, date1904);
        if (stored === null) continue;
        const value = zero ? 0 : arrayAt(result, r - range.top, c - range.left);
        if (!sameValue(stored, value === null ? 0 : value)) return true;
      }
    }
    return false;
  }

  /** `rec` was asked for while it was being calculated: every formula waiting on it is part of a circle. */
  private markLoop(rec: FormulaRec): void {
    const chain = this.chain;
    for (let k = rec.position; k < chain.length; k++) {
      const member = chain[k];
      if (member.looped) continue;
      member.looped = true;
      this.loops.push(member);
    }
  }

  // -------------------------------------------------------------------------
  // Writing results

  /** Writes the results of the formulas that were finished since the last time. */
  private flush(): void {
    const done = this.done;
    if (done.length === 0) return;
    // What spills left is emptied now (see valueOf()).
    this.vacated = [];
    for (let k = 0; k < done.length; k++) {
      this.store(done[k]);
      // One taken up again since (see reopen()) is still to be calculated, with what it reads changed.
      if (done[k].state === DONE) done[k].live = false;
    }
    done.length = 0;
  }

  private store(rec: FormulaRec): void {
    // What a recalculation that ran out of budget did not reach keeps its value.
    if (rec.gone || rec.state !== DONE || rec.keep === KEEP_STORED) return;
    const sh = this.sheets[rec.sheet];
    const modelRow = sh.model.rows.get(rec.row);
    const cell = modelRow ? modelRow.cells[rec.col] : undefined;
    if (!cell || cell.formula !== rec.formula) return;
    // Where the formula spilled before, as the model says (its `ref`): what of it the spill leaves is emptied.
    const before = rec.dynamic ? spillAreaOf(rec.formula, rec.row, rec.col) : undefined;
    if (rec.keep === KEEP_NONE) {
      const placed = this.put(sh, rec.row, rec.col, cell, rec.value);
      if (rec.dynamic) this.placeRef(sh, rec, placed);
    }
    const range = rec.spill;
    if (before) this.clearSpill(rec, before, range);
    if (rec.left) {
      if (!before || !sameRange(before, rec.left)) this.clearSpill(rec, rec.left, range);
      rec.left = undefined;
    }
    if (!range) return;
    const result = rec.result || ZERO;
    const zero = rec.circular || rec.looped;
    const write = (r: number, c: number, existing: Cell | undefined) => {
      if (r === rec.row && c === rec.col) return;
      // A formula of its own is not what the array formula left there, nor is what was typed into the cells of a spill.
      if (existing && (existing.formula || (rec.dynamic && existing.spilled !== true && !isVacant(existing)))) return;
      // Where the result does not reach, arrayAt() answers #N/A, as Excel fills such cells.
      const value = zero ? 0 : arrayAt(result, r - range.top, c - range.left);
      this.put(sh, r, c, existing, value, rec.dynamic, false);
    };
    const size = spillSize(range);
    // A dynamic one spills no more than its share (see spillShare()), or the cells the file has. The cells of an array
    // formula of the file are as many as its `ref` says, which a damaged file can make millions of: those the file
    // does not have are made up to a budget per recalculation (what reads them gets the formula's values all the
    // same, see forEachCell()), past it only those there are get their values.
    if (rec.dynamic || size <= this.fillLeft) {
      if (!rec.dynamic) this.fillLeft -= size;
      for (let r = range.top; r <= range.bottom; r++) {
        const row = sh.model.rows.get(r);
        for (let c = range.left; c <= range.right; c++) write(r, c, row ? row.cells[c] : undefined);
      }
    } else this.eachCell(sh, range, (cell, r, c) => write(r, c, cell));
  }

  /**
   * Writes where a dynamic array formula spills into its `ref`, the way files have it: the
   * cells it fills, or its own cell alone (one value, or #SPILL!).
   */
  private placeRef(sh: SheetState, rec: FormulaRec, cell: Cell): void {
    const formula = rec.formula;
    const own = rangeName({ top: rec.row, left: rec.col, bottom: rec.row, right: rec.col });
    const ref = rec.spill ? rangeName(rec.spill) : own;
    if (ref === formula.ref) return;
    const next: CellFormula = { ...formula, ref };
    const replaced: Cell = { ...cell, formula: next };
    this.own.add(replaced);
    const row = sh.model.rows.get(rec.row) as Row;
    if (!sh.copied.has(rec.row)) {
      sh.model.rows.set(rec.row, copyRow(row, row.cells.slice()));
      sh.copied.add(rec.row);
    }
    (sh.model.rows.get(rec.row) as Row).cells[rec.col] = replaced;
    (sh.model.changedRows || (sh.model.changedRows = new Set())).add(rec.row);
    sh.model.modified = true;
    touchModel(this.workbook);
    // The record stays the formula's: the cell holds what it was, with where its result goes.
    rec.formula = next;
  }

  /**
   * Puts a value into a cell by replacing the cell and its row, when it differs from what is
   * there. `spilled`: the value is what a dynamic array formula spills there (Cell.spilled).
   * `own`: the cell goes into the set of cells an engine wrote (see OWN). The cells an array
   * formula fills do not: the set is asked of them only while the formula's own cell holds the
   * value of the file, and then the engine wrote none of them; and a set of millions of cells
   * costs seconds to fill.
   */
  private put(sh: SheetState, row: number, col: number, cell: Cell | undefined, value: Scalar, spilled = false, own = true): Cell {
    let kind: Cell['kind'];
    let stored: number | string | boolean;
    if (value === null || value === undefined) (kind = 'number'), (stored = 0);
    else if (typeof value === 'number') (kind = 'number'), (stored = isFinite(value) ? value : 0);
    else if (typeof value === 'string') (kind = 'string'), (stored = value);
    else if (typeof value === 'boolean') (kind = 'boolean'), (stored = value);
    else if (isError(value)) (kind = 'error'), (stored = value.error);
    else (kind = 'error'), (stored = '#VALUE!');
    if (cell && cell.kind === kind && cell.value === stored && (!spilled || cell.spilled === true)) return cell;
    const next: Cell = cell ? { ...cell, kind, value: stored } : { kind, value: stored, style: 0 };
    // The text no longer comes from the shared string the cell was read from.
    delete next.runs;
    delete next.si;
    if (spilled) next.spilled = true;
    if (own) this.own.add(next);
    // The cells of an array formula that is not a dynamic one hold what it gives (they are no spill's): how far their
    // column holds something follows them.
    const holds = CalcEngine.holds(next);
    if (holds !== CalcEngine.holds(cell)) {
      sh.contentVersion++;
      if (sh.lastRows) {
        const last = sh.lastRows.get(col) || 0;
        if (holds) {
          if (row + 1 > last) sh.lastRows.set(col, row + 1);
        } else if (row + 1 >= last && last > 0) sh.unsure.add(col);
      }
    }
    // A result that is written is a change of the workbook like any other: a file that is being written while
    // results arrive (a calculation that runs in slices) would hold some of before and some of after.
    touchModel(this.workbook);
    const model = sh.model;
    let target = model.rows.get(row);
    if (!target) {
      target = { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] };
      MADE_ROWS.add(target);
      model.rows.set(row, target);
      sh.copied.add(row);
      sh.rowKeys = undefined;
    } else if (!sh.copied.has(row)) {
      target = copyRow(target, target.cells.slice());
      model.rows.set(row, target);
      sh.copied.add(row);
    }
    target.cells[col] = next;
    (model.changedRows || (model.changedRows = new Set())).add(row);
    model.modified = true;
    if (row >= model.rowCount) model.rowCount = row + 1;
    if (col >= model.colCount) model.colCount = col + 1;
    sh.changed.push({ row, col });
    return next;
  }

  // -------------------------------------------------------------------------
  // The dependency graph

  /** Sheets and names: all that evaluating one formula needs. */
  private prepare(sheetsOfBefore = false): void {
    if (this.prepared) {
      const gone = sheetsOfBefore ? undefined : this.sheetsChanged();
      if (gone === undefined) return;
      // The graph knows sheets by their places and names: it is of a workbook that is no more.
      // What a calculation under way had left, and what read a sheet that went, cannot be told
      // apart from the rest any more: everything is calculated, and the values of the file give
      // way. An engine that has no graph yet knows nothing that could be lost with it.
      const all = this.incomplete || (gone && this.built);
      this.invalidate();
      if (all) {
        this.wantAll = true;
        this.wantOverwrite = true;
        this.incomplete = true;
      }
    }
    this.prepared = true;
    const workbook = this.workbook;
    this.sheets = (workbook.sheets || []).map((model, index) => ({
      index,
      model,
      name: model.name,
      recs: new Map(),
      cellDeps: new Map(),
      rowBands: new Map(),
      colBands: new Map(),
      wide: { entries: [], done: 0, doneLive: 0 },
      spillCells: new Map(),
      spillRanges: [],
      shadowed: new Map(),
      spills: new Set(),
      ranges: new Map(),
      columnKeys: [],
      forcedRecs: new Set(),
      foreign: new Map(),
      unsure: new Set(),
      contentVersion: 0,
      wideRanges: new Map(),
      spillCols: new Map(),
      wideSpills: new Set(),
      rowEpoch: new Uint32Array(Math.min(Math.max(model.rowCount || 0, 0), MAX_ROWS) + 16),
      copied: new Set(),
      changed: [],
    }));
    this.sheetByName = new Map();
    this.sheets.forEach((sh) => {
      const key = foldName(sh.model.name);
      if (!this.sheetByName.has(key)) this.sheetByName.set(key, sh.index);
    });
    this.readNames(undefined);
  }

  /**
   * Whether the list of sheets is no longer the one that was read: undefined when it is;
   * true when a sheet went that formulas of the graph read, or may have read (a formula that
   * chooses its cells while calculating reads any sheet); false when sheets only came, moved
   * or got other names, or the ones that went were read by nothing.
   */
  private sheetsChanged(): boolean | undefined {
    const now = this.workbook.sheets || [];
    const read = this.sheets;
    let same = now.length === read.length;
    for (let k = 0; same && k < read.length; k++) same = read[k].model === now[k] && read[k].name === now[k].name;
    if (same) return undefined;
    for (let k = 0; k < read.length; k++) {
      const sh = read[k];
      if (now.indexOf(sh.model) >= 0) continue;
      if (this.volatiles.size > 0 || this.readFromElsewhere(sh)) return true;
    }
    return false;
  }

  /** Whether a formula of another sheet reads cells of this one: the formulas of a sheet that goes go with it. */
  private readFromElsewhere(sh: SheetState): boolean {
    const index = sh.index;
    for (const list of sh.cellDeps.values()) for (const rec of list) if (rec.sheet !== index) return true;
    const buckets: Bucket[] = [sh.wide];
    sh.rowBands.forEach((bucket) => buckets.push(bucket));
    sh.colBands.forEach((bucket) => buckets.push(bucket));
    for (const bucket of buckets) for (const entry of bucket.entries) if (entry.rec.sheet !== index) return true;
    return false;
  }

  /**
   * Reads the defined names of the workbook. With the names as they were, gives the names
   * (in upper case) that are not what they were, in any scope; what was parsed of the others is kept.
   */
  private readNames(before: Map<string, { global?: NameEntry; local?: Map<number, NameEntry> }> | undefined): Set<string> {
    const names = new Map<string, { global?: NameEntry; local?: Map<number, NameEntry> }>();
    const changed = new Set<string>();
    for (const name of this.workbook.definedNames || []) {
      if (!name || typeof name.name !== 'string' || typeof name.value !== 'string') continue;
      const key = name.name.toUpperCase();
      let found = names.get(key);
      if (!found) names.set(key, (found = {}));
      const old = before ? before.get(key) : undefined;
      const local = name.localSheet === undefined ? undefined : name.localSheet;
      if (local === undefined ? found.global !== undefined : found.local !== undefined && found.local.has(local)) continue;
      const was = old ? (local === undefined ? old.global : old.local ? old.local.get(local) : undefined) : undefined;
      const entry: NameEntry = was && was.text === name.value ? was : { text: name.value };
      if (entry !== was) changed.add(key);
      if (local === undefined) found.global = entry;
      else (found.local || (found.local = new Map())).set(local, entry);
    }
    if (before) {
      // Names, or scopes of a name, that are no longer there.
      before.forEach((old, key) => {
        const found = names.get(key);
        if (!found || (old.global !== undefined && found.global === undefined) || (old.local ? old.local.size : 0) !== (found.local ? found.local.size : 0)) changed.add(key);
      });
    }
    this.names = names;
    return changed;
  }

  /**
   * Reads the formulas of the workbook into the graph, as far as the time of this call
   * goes; false when there are more. Every call reads some.
   */
  private build(): boolean {
    if (this.built) return true;
    const state = this.building || (this.building = { sheet: 0, masters: {} });
    let count = 0;
    for (; state.sheet < this.sheets.length; state.sheet++, state.rows = undefined, state.masters = {}) {
      const sh = this.sheets[state.sheet];
      if (!state.rows) state.rows = sh.model.rows.entries();
      for (;;) {
        const next = state.rows.next();
        if (next.done) break;
        const index = next.value[0];
        const row = next.value[1];
        if (!isIndex(index, MAX_ROWS) || !row || !row.cells) continue;
        const cells = row.cells;
        for (let col = 0; col < cells.length && col < MAX_COLS; col++) {
          const cell = cells[col];
          if (!cell || !cell.formula) continue;
          this.addFormula(sh, index, col, cell, state.masters);
          count++;
        }
        if (count >= BUILD_EVERY) {
          count = 0;
          if (this.outOfTime()) return false;
        }
      }
    }
    this.built = true;
    this.building = undefined;
    this.searchAll = true;
    return true;
  }

  /** Reads the formula of a cell into the graph. */
  private addFormula(sh: SheetState, row: number, col: number, cell: Cell, holder: MastersHolder): void {
    const formula = cell.formula as CellFormula;
    // The inputs of a data table are substituted by Excel itself; its cells keep their values.
    if (formula.type === 'dataTable') {
      this.looked.add(cell);
      return;
    }
    let text = formula.text;
    let dRow = 0;
    let dCol = 0;
    if (formula.type === 'shared' && text === '') {
      const master = formula.si === undefined ? undefined : (holder.masters || (holder.masters = sharedMasters(sh.model))).get(formula.si);
      // Without its master the cell cannot be calculated; it keeps its value.
      if (!master) {
        this.looked.add(cell);
        this.noteForeign(sh, row, col, formula);
        return;
      }
      text = master.text;
      dRow = row - master.row;
      dCol = col - master.col;
    }
    let node: Node;
    try {
      node = parseFormula(text);
    } catch (error) {
      this.unparsedCells.set(`${sh.index}:${row}:${col}`, { sheet: sh.index, row, col });
      this.looked.add(cell);
      this.noteForeign(sh, row, col, formula);
      return;
    }
    // What cannot be calculated here keeps the value the file has for it. Without such a value
    // (a formula just typed in, an error the engine wrote itself) it is calculated anyway, so
    // that the cell shows the error.
    if (hasValue(cell) && !this.own.has(cell)) {
      this.metNames = [];
      if (!this.supports(node, sh.index)) {
        this.unsupportedCells.set(`${sh.index}:${row}:${col}`, { sheet: sh.index, row, col });
        // A name it lacks may be defined later (see namesChanged()).
        if (this.metNames.length > 0) this.unsupportedNames.set(`${sh.index}:${row}:${col}`, this.metNames);
        this.looked.add(cell);
        this.noteForeign(sh, row, col, formula);
        return;
      }
    }
    const rec: FormulaRec = {
      sheet: sh.index,
      row,
      col,
      formula,
      text,
      node,
      dRow,
      dCol,
      volatile: false,
      timed: false,
      isArray: formula.type === 'array' && formula.dynamic !== true,
      // Excel marks the cell of a dynamic array formula with cell metadata.
      arrayMode: formula.type === 'array' || cell.cm !== undefined || formula.dynamic === true,
      dynamic: formula.dynamic === true && formula.type !== 'shared',
      reopens: 0,
      reopenEpoch: 0,
      selfRef: false,
      rowAware: false,
      costly: false,
      precedents: [],
      epoch: 0,
      state: FRESH,
      value: null,
      circular: false,
      looped: false,
      live: false,
      keep: KEEP_NONE,
      kept: false,
      gone: false,
      position: 0,
      stamp: 0,
      index: 0,
      low: 0,
      onStack: false,
      comp: 0,
      at: 0,
    };
    if ((rec.isArray || rec.dynamic) && typeof formula.ref === 'string') {
      const range = parseRange(formula.ref);
      // A dynamic array formula spills down and to the right of its cell.
      const inside = range && (rec.dynamic ? row === range.top && col === range.left : row >= range.top && row <= range.bottom && col >= range.left && col <= range.right);
      if (range && inside) {
        const cells = (range.bottom - range.top + 1) * (range.right - range.left + 1);
        if (cells > 1 && cells <= MAX_SPILL_CELLS) this.registerSpill(sh, rec, range);
      }
    }
    // A dynamic array formula that fills its own cell alone may want more (what is in the way, or a circle, is not in
    // the file): it is calculated once, by the next recalculation, to know where it would spill. Its value stays.
    if (rec.dynamic && !rec.spill && hasValue(cell)) this.unguarded.push(rec);
    // Which cells are its spill's the model says (Cell.spilled): nothing of them is kept here.
    if (rec.dynamic) this.setGuard(rec, rec.spill || { top: row, left: col, bottom: row, right: col });
    this.findPrecedents(rec);
    rec.circular = rec.selfRef;
    if (rec.volatile) this.volatiles.add(rec);
    if (rec.rowAware) this.rowAware.add(rec);
    if (rec.selfRef) this.selfRefs.add(rec);
    if (rec.names) this.named.add(rec);
    sh.recs.set(keyOf(row, col), rec);
    this.graphChanged(sh);
    // While the graph is being built every formula is looked at for circular references in the end.
    if (this.built) this.seeds.push(rec);
  }

  /** An array formula left out of the graph: the cells its `ref` names stay its (see SheetState.foreign). */
  private noteForeign(sh: SheetState, row: number, col: number, formula: CellFormula): void {
    if (!(formula.type === 'array' || formula.dynamic === true) || typeof formula.ref !== 'string') return;
    const range = parseRange(formula.ref);
    if (!range || !(row >= range.top && row <= range.bottom && col >= range.left && col <= range.right)) return;
    if (range.top === range.bottom && range.left === range.right) return;
    if (spillSize(range) > MAX_SPILL_CELLS) return;
    sh.foreign.set(keyOf(row, col), range);
  }

  /** Whether a cell is one an array formula outside the graph fills. */
  private isForeign(sh: SheetState, row: number, col: number): boolean {
    if (sh.foreign.size === 0) return false;
    for (const range of sh.foreign.values()) {
      if (row >= range.top && row <= range.bottom && col >= range.left && col <= range.right) return true;
    }
    return false;
  }

  /**
   * Whether everything a formula uses can be calculated here: no function the library lacks,
   * no LET or LAMBDA, no reference across sheets or into another workbook, no sheet or name
   * the workbook does not have (the name of a table is one: tables are not in the model), no
   * defined name that cannot be parsed or uses one of those, and no use of a function that
   * the evaluator only knows in part.
   */
  private supports(root: Node, sheet: number): boolean {
    const names = new Set<string>();
    // Names of LET and parameters of LAMBDA, in the formula and in the defined names it uses.
    const bound = new Set<string>();
    const binds = (tree: Node) => {
      const found = boundNames(tree);
      if (found) found.forEach((name) => bound.add(name));
    };
    binds(root);
    const nodes: Node[] = [root];
    while (nodes.length > 0) {
      const node = nodes.pop() as Node;
      switch (node.type) {
        case 'ref':
          if (node.external || node.sheetEnd !== undefined) return false;
          if (node.sheet !== undefined && this.sheetIndex(node.sheet) === undefined) return false;
          break;
        case 'name': {
          if (node.sheet === undefined && bound.has(node.name.toUpperCase())) break;
          this.metNames.push(node.name.toUpperCase());
          const home = node.sheet === undefined ? sheet : this.sheetIndex(node.sheet);
          if (home === undefined) return false;
          const key = `${home}!${node.name.toUpperCase()}`;
          if (names.has(key)) break;
          names.add(key);
          const target = this.nameTree(node.name, home);
          if (!target) return false;
          binds(target);
          nodes.push(target);
          break;
        }
        case 'call': {
          const builtin = isBuiltinForm(node.name);
          const spec = builtin ? undefined : this.functions.get(node.name);
          if (!builtin && !spec) {
            // A LAMBDA of the formula, or a defined name that stands for one.
            const upper = node.name.toUpperCase();
            if (!bound.has(upper)) {
              this.metNames.push(upper);
              const key = `${sheet}!${upper}`;
              if (!names.has(key)) {
                names.add(key);
                const target = this.nameTree(node.name, sheet);
                if (!target) return false;
                binds(target);
                nodes.push(target);
              }
            }
            for (const arg of node.args) if (arg) nodes.push(arg);
            break;
          }
          if (!this.supportsCall(node.name, node.args)) return false;
          // A union of references is taken apart by AREAS, by functions of any number of arguments
          // and by INDEX, which picks one of the areas of its first argument.
          const takesUnion = node.name === 'AREAS' || (spec !== undefined && spec.max === undefined);
          for (let k = 0; k < node.args.length; k++) {
            const arg = node.args[k];
            if (!arg) continue;
            if (takesUnion || (k === 0 && node.name === 'INDEX')) {
              const parts = unionParts(arg);
              for (let j = 0; j < parts.length; j++) nodes.push(parts[j]);
            } else nodes.push(arg);
          }
          break;
        }
        case 'unary':
          nodes.push(node.operand);
          break;
        case 'binary':
          // A union anywhere else has no value here.
          if (node.op === ',') return false;
          nodes.push(node.left, node.right);
          break;
      }
    }
    return true;
  }

  /** The uses of functions the evaluator knows only in part. */
  private supportsCall(name: string, args: (Node | null)[]): boolean {
    switch (name) {
      case 'CELL': {
        // What is asked for must be known before the formula is calculated.
        const info = args[0];
        return !!info && info.type === 'string' && CELL_INFO.has(info.value.toLowerCase());
      }
      case 'INDIRECT': {
        // R1C1 notation is not supported: the second argument must say "A1" for certain.
        const a1 = args.length > 1 ? args[1] : undefined;
        if (!a1) return true;
        return (a1.type === 'boolean' && a1.value) || (a1.type === 'number' && a1.value !== 0);
      }
      case 'TEXT': {
        // A format that is not a constant is calculated: what it holds is known only then.
        const format = args.length > 1 ? args[1] : undefined;
        return !format || format.type !== 'string' || !foreignFormat(format.value);
      }
    }
    return true;
  }

  /** Brings the graph in line with a cell that may have been changed. */
  private refresh(sh: SheetState, row: number, col: number): void {
    const key = keyOf(row, col);
    const rec = sh.recs.get(key);
    const modelRow = sh.model.rows.get(row);
    const cell = modelRow ? modelRow.cells[col] : undefined;
    const formula = cell ? cell.formula : undefined;
    if (rec && rec.formula === formula) return;
    this.reread(sh, row, col);
  }

  /** Reads the formula of a cell into the graph again, whatever was known about the cell. */
  private reread(sh: SheetState, row: number, col: number): void {
    const rec = sh.recs.get(keyOf(row, col));
    const modelRow = sh.model.rows.get(row);
    const cell = modelRow ? modelRow.cells[col] : undefined;
    const name = `${sh.index}:${row}:${col}`;
    this.unparsedCells.delete(name);
    this.unsupportedCells.delete(name);
    this.unsupportedNames.delete(name);
    this.mismatchedCells.delete(name);
    sh.foreign.delete(keyOf(row, col));
    if (this.costlyCells.delete(name)) this.staleCells = undefined;
    if (rec) this.removeFormula(sh, rec);
    if (cell && cell.formula) this.addFormula(sh, row, col, cell, {});
  }

  /** The formulas of a sheet changed: what was worked out from them is worked out again. */
  private graphChanged(sh: SheetState): void {
    sh.columns = undefined;
    this.staleCells = undefined;
    const search = this.search;
    if (!search) return;
    // The search in progress walked a graph that is no longer there: it starts over.
    if (search.all) this.searchAll = true;
    else for (const rec of search.seeds) this.seeds.push(rec);
    this.search = undefined;
  }

  private removeFormula(sh: SheetState, rec: FormulaRec): void {
    sh.recs.delete(keyOf(rec.row, rec.col));
    sh.spills.delete(rec);
    this.volatiles.delete(rec);
    this.rowAware.delete(rec);
    this.selfRefs.delete(rec);
    this.named.delete(rec);
    // Whoever still holds the record (the list of dirty formulas) must not write it.
    rec.gone = true;
    this.graphChanged(sh);
    if (rec.comp) {
      // The formulas it made a circle with may no longer make one.
      const members = this.groups.get(rec.comp);
      this.groups.delete(rec.comp);
      if (members) {
        for (const member of members) {
          member.comp = 0;
          this.setCircular(member, member.selfRef);
          if (member !== rec) this.seeds.push(member);
        }
      }
    }
    this.budgetRefused.delete(rec);
    // The formulas found circular with it are settled again without it.
    if (rec.forcedGroup) this.clearCircle(rec);
    this.unregisterSpill(sh, rec);
    // What a dynamic array formula that went had filled is emptied, and what reads it is calculated again.
    const area = rec.dynamic ? spillAreaOf(rec.formula, rec.row, rec.col) : undefined;
    if (area && this.clearSpill(rec, area, undefined)) this.cleared.push({ sheet: sh.index, range: area });
    if (rec.left && !sameRange(rec.left, area) && this.clearSpill(rec, rec.left, undefined)) this.cleared.push({ sheet: sh.index, range: rec.left });
    rec.left = undefined;
    for (const precedent of rec.precedents) {
      this.unindex(rec, precedent);
      const target = this.sheets[precedent.sheet];
      const wide = target && !precedent.shapeOnly && isWholeLine(precedent.range) ? target.wideRanges.get(rangeKey(precedent.range)) : undefined;
      if (wide) {
        wide.readers.delete(rec);
        if (wide.readers.size === 0) target.wideRanges.delete(rangeKey(precedent.range));
      }
    }
    rec.precedents = [];
    rec.guard = undefined;
    sh.forcedRecs.delete(rec);
  }

  /** Takes one precedent of a formula out of the index of the sheet it is on. */
  private unindex(rec: FormulaRec, precedent: Precedent): void {
    const target = this.sheets[precedent.sheet];
    if (!target) return;
    const range = precedent.range;
    if (range.top === range.bottom && range.left === range.right) {
      const key = keyOf(range.top, range.left);
      const list = target.cellDeps.get(key);
      if (list) {
        // The same cell may be read twice, once only as a guard; one entry goes.
        const at = list.indexOf(rec);
        const kept = at < 0 ? list : list.slice(0, at).concat(list.slice(at + 1));
        if (kept.length > 0) target.cellDeps.set(key, kept);
        else target.cellDeps.delete(key);
      }
      return;
    }
    const drop = (bucket: Bucket | undefined) => {
      if (!bucket) return;
      const at = bucket.entries.findIndex((entry) => entry.rec === rec && entry.top === range.top && entry.left === range.left && entry.bottom === range.bottom && entry.right === range.right);
      if (at >= 0) bucket.entries.splice(at, 1);
    };
    const firstBand = range.top >> BAND_SHIFT;
    const lastBand = range.bottom >> BAND_SHIFT;
    if (lastBand - firstBand < MAX_BANDS) for (let band = firstBand; band <= lastBand; band++) drop(target.rowBands.get(band));
    else if (range.right - range.left < MAX_BANDS) for (let c = range.left; c <= range.right; c++) drop(target.colBands.get(c));
    else drop(target.wide);
  }

  /** Makes a range the cells a formula fills, for the formulas that read them to find. */
  private registerSpill(sh: SheetState, rec: FormulaRec, range: CellRange): void {
    rec.spill = range;
    if (rec.dynamic) this.wanting.add(rec);
    sh.spills.add(rec);
    if (range.right - range.left < 64) {
      for (let c = range.left; c <= range.right; c++) {
        let list = sh.spillCols.get(c);
        if (!list) sh.spillCols.set(c, (list = new Set()));
        list.add(rec);
      }
    } else sh.wideSpills.add(rec);
    const cells = (range.bottom - range.top + 1) * (range.right - range.left + 1);
    if (cells <= MAX_LISTED_SPILL) {
      for (let r = range.top; r <= range.bottom; r++) {
        for (let c = range.left; c <= range.right; c++) {
          const key = keyOf(r, c);
          const before = sh.spillCells.get(key);
          if (before && before !== rec) {
            const under = sh.shadowed.get(key);
            if (under) under.push(before);
            else sh.shadowed.set(key, [before]);
          }
          sh.spillCells.set(key, rec);
        }
      }
    } else sh.spillRanges.push(rec);
  }

  private unregisterSpill(sh: SheetState, rec: FormulaRec): void {
    sh.spills.delete(rec);
    const range = rec.spill;
    if (!range) return;
    if (range.right - range.left < 64) {
      for (let c = range.left; c <= range.right; c++) {
        const list = sh.spillCols.get(c);
        if (list && list.delete(rec) && list.size === 0) sh.spillCols.delete(c);
      }
    } else sh.wideSpills.delete(rec);
    if (rec.dynamic && this.vacated.length < MAX_VACATED) this.vacated.push({ sheet: sh.index, range });
    // What it filled since its result was last written (cells made for it, see materialize()) is emptied when it is.
    if (rec.dynamic) rec.left = rec.left ? unionOf(rec.left, range) : range;
    const at = sh.spillRanges.indexOf(rec);
    if (at >= 0) sh.spillRanges.splice(at, 1);
    else {
      for (let r = range.top; r <= range.bottom; r++) {
        for (let c = range.left; c <= range.right; c++) {
          const key = keyOf(r, c);
          const under = sh.shadowed.size > 0 ? sh.shadowed.get(key) : undefined;
          if (sh.spillCells.get(key) === rec) {
            sh.spillCells.delete(key);
            // Another spill registered under it has the cell back.
            while (under && under.length > 0) {
              const other = under.pop() as FormulaRec;
              const area = other.spill;
              if (other !== rec && area && r >= area.top && r <= area.bottom && c >= area.left && c <= area.right) {
                sh.spillCells.set(key, other);
                break;
              }
            }
          } else if (under) {
            const at = under.indexOf(rec);
            if (at >= 0) under.splice(at, 1);
          }
          if (under && under.length === 0) sh.shadowed.delete(key);
        }
      }
    }
    rec.spill = undefined;
  }

  /** The cells a dynamic array formula watches for what may come in the way of its result. */
  private setGuard(rec: FormulaRec, area: CellRange): void {
    if (rec.guard) {
      if (sameRange(rec.guard.range, area)) return;
      // Circles run through the cells it would spill into (see successor()): they are looked for again.
      if (spillSize(rec.guard.range) > 1 || spillSize(area) > 1) {
        this.sheets[rec.sheet].columns = undefined;
        this.seeds.push(rec);
      }
      this.unindex(rec, rec.guard);
      const at = rec.precedents.indexOf(rec.guard);
      if (at >= 0) rec.precedents.splice(at, 1);
    }
    this.addPrecedent(rec, rec.sheet, area, true);
    rec.guard = rec.precedents[rec.precedents.length - 1];
  }

  /** Calls `visit` for every cell of a range that exists, row by row. */
  private eachCell(sh: SheetState, area: CellRange, visit: (cell: Cell, row: number, col: number) => void): void {
    const rows = sh.model.rows;
    const each = (index: number) => {
      const row = rows.get(index);
      if (!row || !row.cells) return;
      const last = Math.min(area.right, row.cells.length - 1);
      for (let c = area.left; c <= last; c++) {
        const cell = row.cells[c];
        if (cell) visit(cell, index, c);
      }
    };
    if (area.bottom - area.top < DIRECT_ROWS) for (let r = area.top; r <= area.bottom; r++) each(r);
    else {
      const keys = Array.from(rows.keys()).filter((key) => key >= area.top && key <= area.bottom);
      keys.sort((a, b) => a - b);
      for (const key of keys) each(key);
    }
  }

  /**
   * Empties the cells of `area` outside `keep` that hold what a dynamic array formula spilled
   * there (Cell.spilled): such a cell keeps its format, a cell with nothing else goes, and so
   * does a row that holds nothing then. A cell an edit put there since is left alone. Whether
   * anything changed.
   */
  private clearSpill(rec: FormulaRec, area: CellRange, keep: CellRange | undefined): boolean {
    const sh = this.sheets[rec.sheet];
    const model = sh.model;
    // By row: each row is copied once, however many of its cells change.
    const byRow = new Map<number, number[]>();
    this.eachCell(sh, area, (cell, r, c) => {
      if (r === rec.row && c === rec.col) return;
      if (keep && r >= keep.top && r <= keep.bottom && c >= keep.left && c <= keep.right) return;
      if (cell.spilled !== true || cell.formula) return;
      // Another formula spills there now: the cell is its.
      const owner = sh.spillCells.size > 0 || sh.spillRanges.length > 0 ? this.spillOwner(sh, r, c) : undefined;
      if (owner && owner !== rec && !owner.gone) return;
      const list = byRow.get(r);
      if (list) list.push(c);
      else byRow.set(r, [c]);
    });
    if (byRow.size === 0) return false;
    byRow.forEach((cols, r) => {
      const row = model.rows.get(r);
      if (!row) return;
      const cells = row.cells.slice();
      for (const c of cols) {
        const cell = cells[c];
        if (cell.style) cells[c] = { kind: 'blank', value: null, style: cell.style };
        else delete cells[c];
        sh.changed.push({ row: r, col: c });
      }
      while (cells.length > 0 && !(cells.length - 1 in cells)) cells.length--;
      const next: Row = copyRow(row, cells);
      // A row the engine made for the spill goes when it holds nothing; one the workbook had stays, empty or not.
      const plain = MADE_ROWS.has(row) && cells.length === 0 && !next.height && !next.customHeight && !next.hidden && !next.outlineLevel && !next.collapsed && next.style === undefined && !next.attrs;
      if (plain) {
        model.rows.delete(r);
        sh.rowKeys = undefined;
      } else model.rows.set(r, next);
      sh.copied.add(r);
      (model.changedRows || (model.changedRows = new Set())).add(r);
    });
    model.modified = true;
    touchModel(this.workbook);
    return true;
  }

  /** Empties a cell marked as a spill's that no spill fills (see recalculate()); it keeps its format. */
  private dropStray(sh: SheetState, r: number, c: number): void {
    const model = sh.model;
    const row = model.rows.get(r);
    const cell = row ? row.cells[c] : undefined;
    if (!row || !cell || cell.spilled !== true || cell.formula) return;
    if ((sh.spillCells.size > 0 || sh.spillRanges.length > 0) && this.spillOwner(sh, r, c)) return;
    if (this.isForeign(sh, r, c)) return;
    const cells = row.cells.slice();
    if (cell.style) cells[c] = { kind: 'blank', value: null, style: cell.style };
    else delete cells[c];
    while (cells.length > 0 && !(cells.length - 1 in cells)) cells.length--;
    model.rows.set(r, copyRow(row, cells));
    sh.copied.add(r);
    sh.changed.push({ row: r, col: c });
    (model.changedRows || (model.changedRows = new Set())).add(r);
    model.modified = true;
    touchModel(this.workbook);
  }

  /** Walks the tree of a formula for the cells it reads. */
  private findPrecedents(rec: FormulaRec): void {
    const seen = new Set<string>();
    const names = new Set<string>();
    // The tree can be thousands of operators deep on one side; it is walked with a stack of its own.
    const nodes: Node[] = [rec.node];
    // Whether the node at the same index comes from the formula itself (and so moves with a shared formula).
    const own: boolean[] = [true];
    // Whether it is a reference given to a function that only asks where it is.
    const shape: boolean[] = [false];
    // Whether it is a reference a function picks cells from (INDEX): which of them are read is known only when calculating.
    const picked: boolean[] = [false];
    // Names of LET and parameters of LAMBDA: no defined names.
    const bound = new Set<string>();
    const binds = (tree: Node) => {
      const found = boundNames(tree);
      if (found) found.forEach((name) => bound.add(name));
    };
    binds(rec.node);
    /** A defined name the formula uses: its tree is walked as part of the formula. */
    const useName = (name: string, sheetName: string | undefined, shapeOnly: boolean, pickedFrom: boolean) => {
      (rec.names || (rec.names = [])).push(name.toUpperCase());
      const sheet = sheetName === undefined ? rec.sheet : this.sheetIndex(sheetName);
      if (sheet === undefined) return;
      const key = `${sheet}!${name.toUpperCase()}:${shapeOnly ? 1 : 0}`;
      if (names.has(key)) return;
      names.add(key);
      const target = this.resolveName(name, sheet);
      if (target) {
        binds(target);
        nodes.push(target);
        own.push(false);
        shape.push(shapeOnly && target.type === 'ref');
        picked.push(pickedFrom && target.type === 'ref');
      }
    };
    while (nodes.length > 0) {
      const node = nodes.pop() as Node;
      const shifted = own.pop() as boolean;
      const shapeOnly = shape.pop() as boolean;
      const pickedFrom = picked.pop() as boolean;
      switch (node.type) {
        case 'ref': {
          if (node.external || node.sheetEnd !== undefined) break;
          const sheet = node.sheet === undefined ? rec.sheet : this.sheetIndex(node.sheet);
          const range = sheet === undefined ? undefined : shifted ? refRange(node, rec.dRow, rec.dCol) : node.range;
          if (sheet === undefined || !range) break;
          // A formula that reads its own cell is circular, whether or not that part of it is ever calculated.
          if (!shapeOnly && !pickedFrom && sheet === rec.sheet && rec.row >= range.top && rec.row <= range.bottom && rec.col >= range.left && rec.col <= range.right) rec.selfRef = true;
          const key = `${sheet}:${range.top}:${range.left}:${range.bottom}:${range.right}:${shapeOnly ? 1 : 0}`;
          if (seen.has(key)) break;
          seen.add(key);
          this.addPrecedent(rec, sheet, range, shapeOnly);
          break;
        }
        case 'name': {
          if (node.sheet === undefined && bound.has(node.name.toUpperCase())) break;
          useName(node.name, node.sheet, shapeOnly, pickedFrom);
          break;
        }
        case 'call': {
          // A call of a LAMBDA a defined name stands for reads what the name reads.
          if (!isBuiltinForm(node.name) && !this.functions.has(node.name) && !bound.has(node.name.toUpperCase())) useName(node.name, undefined, false, false);
          if (VOLATILE.has(node.name)) rec.volatile = true;
          if (TIMED.has(node.name)) rec.timed = true;
          if (ROW_AWARE.has(node.name)) rec.rowAware = true;
          const asksWhere = SHAPE_ONLY.has(node.name);
          // CELL("row", A1) asks where A1 is; CELL("contents", A1) reads it.
          const info = node.name === 'CELL' ? node.args[0] : undefined;
          const cellWhere = !!info && info.type === 'string' && CELL_WHERE.has(info.value.toLowerCase());
          for (let k = 0; k < node.args.length; k++) {
            const arg = node.args[k];
            if (!arg) continue;
            const reference = arg.type === 'ref' || arg.type === 'name';
            nodes.push(arg);
            own.push(shifted);
            shape.push((asksWhere || (k === 0 && node.name === 'OFFSET') || (k === 1 && cellWhere)) && reference);
            picked.push(k === 0 && node.name === 'INDEX' && reference);
          }
          break;
        }
        case 'unary':
          nodes.push(node.operand);
          own.push(shifted);
          shape.push(false);
          picked.push(false);
          break;
        case 'binary':
          // A range put together while calculating (A1:INDEX(...)) can reach cells no reference names.
          if (node.op === ':') rec.volatile = true;
          nodes.push(node.left, node.right);
          own.push(shifted, shifted);
          shape.push(false, false);
          picked.push(false, false);
          break;
      }
    }
  }

  private addPrecedent(rec: FormulaRec, sheet: number, range: CellRange, shapeOnly: boolean): void {
    const target = this.sheets[sheet];
    if (!target) return;
    if (!shapeOnly && isWholeLine(range)) {
      const key = rangeKey(range);
      let wide = target.wideRanges.get(key);
      if (!wide) target.wideRanges.set(key, (wide = { range, readers: new Set() }));
      wide.readers.add(rec);
    }
    const precedent: Precedent = { sheet, range, shapeOnly };
    rec.precedents.push(precedent);
    if (range.top === range.bottom && range.left === range.right) {
      const key = keyOf(range.top, range.left);
      const list = target.cellDeps.get(key);
      if (list) list.push(rec);
      else target.cellDeps.set(key, [rec]);
      return;
    }
    if (!shapeOnly) {
      const name = `${range.top}:${range.left}:${range.bottom}:${range.right}`;
      let node = target.ranges.get(name);
      if (!node) target.ranges.set(name, (node = { area: true, sheet, range, stamp: 0, index: 0, low: 0, onStack: false }));
      precedent.node = node;
    }
    const entry: Entry = { top: range.top, left: range.left, bottom: range.bottom, right: range.right, rec };
    const add = (map: Map<number, Bucket>, key: number) => {
      const bucket = map.get(key);
      if (bucket) {
        bucket.entries.push(entry);
        bucket.done = 0;
        bucket.doneLive = 0;
      } else map.set(key, { entries: [entry], done: 0, doneLive: 0 });
    };
    const firstBand = range.top >> BAND_SHIFT;
    const lastBand = range.bottom >> BAND_SHIFT;
    if (lastBand - firstBand < MAX_BANDS) for (let band = firstBand; band <= lastBand; band++) add(target.rowBands, band);
    else if (range.right - range.left < MAX_BANDS) for (let c = range.left; c <= range.right; c++) add(target.colBands, c);
    else {
      target.wide.entries.push(entry);
      target.wide.done = 0;
      target.wide.doneLive = 0;
    }
  }

  // -------------------------------------------------------------------------
  // Circular references among the references the formulas name
  //
  // The formulas and the ranges they read make a graph: a formula points at every range it
  // reads, a range at every formula inside it. Formulas that can reach one another in it
  // are circular (the strongly connected parts of the graph, found with Tarjan's algorithm).
  // A range is one node however many formulas read it, and a formula whose part is known is
  // struck out of the lists the ranges are read from, so the search looks at every formula
  // and every reference a few times, not at every pair of them. Like building the graph it
  // is done in slices; a change of a formula starts it again from the formulas the change
  // can have tied together or set free.

  /** Goes on finding circular references, as far as the time of this call goes; false when there is more to do. */
  private components(): boolean {
    let run = this.search;
    if (!run) {
      if (!this.searchAll && this.seeds.length === 0) return true;
      let seeds = this.seeds;
      let total = 0;
      for (const sh of this.sheets) total += sh.recs.size;
      if (this.searchAll) {
        seeds = [];
        for (const sh of this.sheets) sh.recs.forEach((rec) => seeds.push(rec));
      }
      run = { all: this.searchAll, seeds, next: 0, stamp: ++this.stamps, counter: 0, stack: [], path: [], a: [], b: [], c: [], d: [], steps: 0, limit: 64 * (total + seeds.length) + 4000000 };
      this.search = run;
      this.searchAll = false;
      this.seeds = [];
      for (const sh of this.sheets) this.indexColumns(sh);
    }
    const path = run.path;
    const stamp = run.stamp;
    let since = 0;
    for (;;) {
      if (path.length === 0) {
        let seed: FormulaRec | undefined;
        while (run.next < run.seeds.length && !seed) {
          const candidate = run.seeds[run.next++];
          if (!candidate.gone && candidate.stamp !== stamp) seed = candidate;
        }
        if (!seed) break;
        this.enter(run, seed);
      }
      if (++since >= SEARCH_EVERY) {
        since = 0;
        run.steps += SEARCH_EVERY;
        // A graph made to be slow (thousands of ranges that all hold one another's formulas):
        // the search is given up, and only circles met while calculating are found.
        if (run.steps > run.limit) break;
        if (this.outOfTime()) return false;
      }
      const top = path.length - 1;
      const node = path[top];
      const next = this.successor(run, node, top);
      if (next) {
        if (next.stamp !== stamp) this.enter(run, next);
        else if (next.onStack && next.index < node.low) node.low = next.index;
        continue;
      }
      path.pop();
      run.a.pop();
      run.b.pop();
      run.c.pop();
      run.d.pop();
      if (top > 0 && node.low < path[top - 1].low) path[top - 1].low = node.low;
      if (node.low === node.index) this.closeGroup(run, node);
    }
    this.search = undefined;
    return true;
  }

  private enter(run: ComponentRun, node: GraphNode): void {
    node.stamp = run.stamp;
    node.index = node.low = run.counter++;
    node.onStack = true;
    run.stack.push(node);
    run.path.push(node);
    // A range starts at the first column it covers; a formula at its first precedent.
    run.a.push((node as RangeNode).area ? lowerBound(this.sheets[node.sheet].columnKeys, (node as RangeNode).range.left) : 0);
    run.b.push((node as RangeNode).area ? -1 : 0);
    run.c.push(0);
    run.d.push(undefined);
  }

  /** The next node `node` points at, or undefined when there are no more. */
  private successor(run: ComponentRun, node: GraphNode, frame: number): GraphNode | undefined {
    if ((node as RangeNode).area) {
      const range = (node as RangeNode).range;
      const sh = this.sheets[node.sheet];
      const keys = sh.columnKeys;
      const columns = sh.columns as Map<number, Column>;
      let ci = run.a[frame];
      let at = run.b[frame];
      while (ci < keys.length && keys[ci] <= range.right) {
        const column = columns.get(keys[ci]) as Column;
        if (at < 0) at = lowerBound(column.rows, range.top);
        at = find(column.skip, at);
        if (at < column.rows.length && column.rows[at] <= range.bottom) {
          run.a[frame] = ci;
          run.b[frame] = at + 1;
          return column.recs[at];
        }
        ci++;
        at = -1;
      }
      run.a[frame] = ci;
      run.b[frame] = -1;
      // Formulas elsewhere that fill cells of the range, or would and are circular for it (see resolve()).
      let fill = run.d[frame];
      if (!fill) run.d[frame] = fill = this.fillersOf(sh, range);
      const si = run.c[frame];
      if (si < fill.length) {
        run.c[frame] = si + 1;
        return fill[si];
      }
      return undefined;
    }
    const precedents = (node as FormulaRec).precedents;
    let pi = run.a[frame];
    while (pi < precedents.length) {
      const precedent = precedents[pi];
      if (precedent.shapeOnly) {
        pi++;
        continue;
      }
      if (precedent.node) {
        run.a[frame] = pi + 1;
        return precedent.node;
      }
      // One cell: the formula in it, then the formula that fills it.
      const sh = this.sheets[precedent.sheet];
      const row = precedent.range.top;
      const col = precedent.range.left;
      const inCell = sh.recs.get(keyOf(row, col));
      if (run.b[frame] === 0) {
        run.b[frame] = 1;
        run.a[frame] = pi;
        if (inCell) return inCell;
      }
      run.b[frame] = 0;
      pi++;
      const owner = sh.spillCells.size > 0 || sh.spillRanges.length > 0 ? this.spillOwner(sh, row, col) : undefined;
      if (owner && owner !== inCell) {
        run.a[frame] = pi;
        return owner;
      }
    }
    run.a[frame] = pi;
    return undefined;
  }

  /**
   * Each formula whose spill meets a range, once: found by the cells of a small range, by the columns of a narrow one
   * (SheetState.spillCols), or among all spills; what is looked at is spent from the budget.
   */
  private eachSpillMeeting(sh: SheetState, range: CellRange, visit: (rec: FormulaRec) => void): void {
    if (sh.spills.size === 0) return;
    let looked = 0;
    let seen: Set<FormulaRec> | undefined;
    const each = (rec: FormulaRec) => {
      looked++;
      if (!rec.spill || rec.gone || !meets(rec.spill, range)) return;
      if (!seen) seen = new Set();
      if (seen.has(rec)) return;
      seen.add(rec);
      visit(rec);
    };
    const cells = (range.bottom - range.top + 1) * (range.right - range.left + 1);
    if (cells <= 1024 && cells < sh.spills.size) {
      for (let r = range.top; r <= range.bottom; r++) {
        for (let c = range.left; c <= range.right; c++) {
          looked++;
          const key = keyOf(r, c);
          const rec = sh.spillCells.get(key);
          if (rec) each(rec);
          // Spills registered over one another while they are settled (see registerSpill()).
          const under = sh.shadowed.size > 0 ? sh.shadowed.get(key) : undefined;
          if (under) for (const other of under) each(other);
        }
      }
      for (const rec of sh.spillRanges) each(rec);
    } else if (range.right - range.left < 64) {
      for (let c = range.left; c <= range.right; c++) {
        const list = sh.spillCols.get(c);
        if (list) list.forEach(each);
      }
      sh.wideSpills.forEach(each);
    } else sh.spills.forEach(each);
    if (this.calculating) this.spend(looked >> 2);
  }

  /** The formulas that fill cells of a range (their spills), or would and are circular for it (see resolve()). */
  private fillersOf(sh: SheetState, range: CellRange): FormulaRec[] {
    const out: FormulaRec[] = [];
    this.eachSpillMeeting(sh, range, (rec) => out.push(rec));
    if (sh.forcedRecs.size > 0) sh.forcedRecs.forEach((rec) => {
      if (rec.forced && !rec.gone && !rec.spill && rec.guard && meets(rec.guard.range, range)) out.push(rec);
    });
    return out;
  }

  /** `node` is the first of a part of the graph whose nodes all reach one another: everything above it on the stack. */
  private closeGroup(run: ComponentRun, node: GraphNode): void {
    const stack = run.stack;
    const members: FormulaRec[] = [];
    for (;;) {
      const member = stack.pop() as GraphNode;
      member.onStack = false;
      if (!(member as RangeNode).area) {
        const rec = member as FormulaRec;
        members.push(rec);
        // Its part is known: ranges read from now on pass over it.
        if (rec.column) rec.column.skip[rec.at] = rec.at + 1;
      }
      if (member === node) break;
    }
    if (members.length > 1) {
      const id = ++this.groupIds;
      for (const rec of members) {
        if (rec.comp) this.groups.delete(rec.comp);
        rec.comp = id;
        this.setCircular(rec, true);
      }
      this.groups.set(id, members);
    } else if (members.length === 1) {
      // Alone: circular only when it names its own cell.
      const rec = members[0];
      if (rec.comp) this.groups.delete(rec.comp);
      rec.comp = 0;
      this.setCircular(rec, rec.selfRef || this.readsOwnSpill(rec));
    }
  }

  /** Whether a dynamic array formula reads cells it spills into itself: a circle of one. */
  private readsOwnSpill(rec: FormulaRec): boolean {
    if (!rec.dynamic) return false;
    const spill = rec.spill;
    if (!spill) return false;
    for (const precedent of rec.precedents) {
      const range = precedent.range;
      if (precedent.shapeOnly || precedent.sheet !== rec.sheet) continue;
      if (range.top <= spill.bottom && range.bottom >= spill.top && range.left <= spill.right && range.right >= spill.left) return true;
    }
    return false;
  }

  /** Whether a formula is circular; one that changes is calculated again (see recalculate()). */
  private setCircular(rec: FormulaRec, circular: boolean): void {
    if (rec.circular === circular) return;
    // Found circular when the spills were settled in this recalculation (see resolve()): it spills nowhere now, so
    // no circle shows through it any more; it stays circular until a recalculation calculates it anew.
    if (!circular && rec.forced && rec.forcedEpoch === this.epoch) return;
    rec.circular = circular;
    // A search of the whole graph (built afresh) finds what the file was saved with: nothing is calculated for it.
    if (!(this.search && this.search.all) && !this.quiet) this.flipped.push(rec);
  }

  /** The formulas of a sheet by column, in row order, with nothing struck out. */
  private indexColumns(sh: SheetState): void {
    let columns = sh.columns;
    if (columns) {
      columns.forEach((column) => {
        const skip = column.skip;
        for (let k = 0; k < skip.length; k++) skip[k] = k;
      });
      return;
    }
    columns = new Map();
    const unsorted = new Set<Column>();
    sh.recs.forEach((rec) => {
      let column = (columns as Map<number, Column>).get(rec.col);
      if (!column) (columns as Map<number, Column>).set(rec.col, (column = { rows: [], recs: [], skip: new Int32Array(0) }));
      if (column.rows.length > 0 && column.rows[column.rows.length - 1] > rec.row) unsorted.add(column);
      column.rows.push(rec.row);
      column.recs.push(rec);
    });
    unsorted.forEach((column) => {
      column.recs.sort((a, b) => a.row - b.row);
      for (let k = 0; k < column.recs.length; k++) column.rows[k] = column.recs[k].row;
    });
    columns.forEach((column) => {
      const skip = new Int32Array(column.rows.length + 1);
      for (let k = 0; k < skip.length; k++) skip[k] = k;
      column.skip = skip;
      for (let k = 0; k < column.recs.length; k++) {
        column.recs[k].column = column;
        column.recs[k].at = k;
      }
    });
    sh.columns = columns;
    sh.columnKeys = Array.from(columns.keys()).sort((a, b) => a - b);
  }

  // -------------------------------------------------------------------------
  // Marking what has to be calculated

  /** Marks a formula dirty for this recalculation. */
  private touch(rec: FormulaRec): void {
    if (rec.epoch === this.epoch) return;
    // Found circular when spills were settled before: calculated anew, its circle is settled again (see resolve()).
    if (rec.forced && rec.forcedEpoch !== this.epoch) {
      this.clearCircle(rec);
      this.seeds.push(rec);
      this.tooOften.add(rec);
    }
    if (rec.costly) {
      // Something it reads changed: it is tried again.
      rec.costly = false;
      this.costlyCells.delete(`${rec.sheet}:${rec.row}:${rec.col}`);
      this.staleCells = undefined;
    }
    rec.epoch = this.epoch;
    rec.state = FRESH;
    rec.keep = KEEP_NONE;
    rec.result = undefined;
    rec.next = undefined;
    this.dirty.push(rec);
    const sh = this.sheets[rec.sheet];
    if (rec.row >= sh.rowEpoch.length) {
      const grown = new Uint32Array(Math.min(Math.max(rec.row + 1, sh.rowEpoch.length * 2), MAX_ROWS));
      grown.set(sh.rowEpoch);
      sh.rowEpoch = grown;
    }
    sh.rowEpoch[rec.row] = this.epoch;
  }

  /**
   * Marks a formula dirty, and live when something it reads really changed; its dependents
   * are marked in turn when it is new to either.
   */
  private mark(rec: FormulaRec, live: boolean): void {
    const gathering = this.gathering;
    if (gathering) {
      if (!gathering.has(rec)) {
        gathering.add(rec);
        this.walk.push(rec);
      }
      return;
    }
    const fresh = rec.epoch !== this.epoch;
    if (fresh) {
      this.touch(rec);
      // A formula without a value gets one: what reads it reads something new.
      if (!live && !rec.live && this.unvalued(rec)) live = true;
    }
    if (live && !rec.live) {
      rec.live = true;
      this.walk.push(rec);
    } else if (fresh) this.walk.push(rec);
  }

  /** Marks the formulas that read the cells of a formula. */
  private dependents(rec: FormulaRec, live: boolean): void {
    const sh = this.sheets[rec.sheet];
    if (rec.spill) this.collect(sh, rec.spill.top, rec.spill.left, rec.spill.bottom, rec.spill.right, live);
    else this.collect(sh, rec.row, rec.col, rec.row, rec.col, live);
  }

  /** Marks every formula that reads a cell of the range. */
  private collect(sh: SheetState, top: number, left: number, bottom: number, right: number, live: boolean): void {
    const deps = sh.cellDeps;
    if (top === bottom && left === right) {
      const list = deps.get(keyOf(top, left));
      if (list) for (let k = 0; k < list.length; k++) this.mark(list[k], live);
    } else if (deps.size > 0) {
      if ((bottom - top + 1) * (right - left + 1) <= deps.size) {
        for (let r = top; r <= bottom; r++) {
          for (let c = left; c <= right; c++) {
            const list = deps.get(keyOf(r, c));
            if (list) for (let k = 0; k < list.length; k++) this.mark(list[k], live);
          }
        }
      } else {
        deps.forEach((list, key) => {
          const r = Math.floor(key / MAX_COLS);
          const c = key - r * MAX_COLS;
          if (r >= top && r <= bottom && c >= left && c <= right) for (let k = 0; k < list.length; k++) this.mark(list[k], live);
        });
      }
    }
    if (sh.rowBands.size > 0) {
      const first = top >> BAND_SHIFT;
      const last = bottom >> BAND_SHIFT;
      if (last - first < sh.rowBands.size) {
        for (let band = first; band <= last; band++) this.scan(sh.rowBands.get(band), top, left, bottom, right, live);
      } else {
        sh.rowBands.forEach((bucket, band) => {
          if (band >= first && band <= last) this.scan(bucket, top, left, bottom, right, live);
        });
      }
    }
    if (sh.colBands.size > 0) {
      if (right - left < sh.colBands.size) {
        for (let c = left; c <= right; c++) this.scan(sh.colBands.get(c), top, left, bottom, right, live);
      } else {
        sh.colBands.forEach((bucket, c) => {
          if (c >= left && c <= right) this.scan(bucket, top, left, bottom, right, live);
        });
      }
    }
    this.scan(sh.wide, top, left, bottom, right, live);
  }

  private scan(bucket: Bucket | undefined, top: number, left: number, bottom: number, right: number, live: boolean): void {
    if (!bucket) return;
    const gathering = this.gathering;
    const epoch = this.epoch;
    if (!gathering && (live ? bucket.doneLive : bucket.done) === epoch) return;
    const entries = bucket.entries;
    let clean = 0;
    for (let k = 0; k < entries.length; k++) {
      const entry = entries[k];
      const rec = entry.rec;
      if (gathering ? gathering.has(rec) : rec.epoch === epoch && (rec.live || !live)) continue;
      if (entry.top <= bottom && entry.bottom >= top && entry.left <= right && entry.right >= left) this.mark(rec, live);
      else clean++;
    }
    // Nothing left to find here until the next recalculation.
    if (!gathering && clean === 0) {
      bucket.done = epoch;
      if (live) bucket.doneLive = epoch;
    }
  }
}
