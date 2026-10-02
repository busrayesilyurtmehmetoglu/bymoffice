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

import { isoToSerial } from '../format/cell-format';
import { MAX_COLS, MAX_ROWS, parseRange, type Cell, type CellFormula, type CellRange, type Row, type Sheet, type Workbook } from '../model/sheet/types';
import { sharedMasters, type SharedMasters } from '../ops/ops';
import { EvaluationAbort, evaluate, evaluationDepth, isBuiltinForm, refRange, toCellValue, type EvaluateOptions } from './eval';
import { FUNCTIONS } from './functions/index';
import { parseFormula, type Node } from './parse';
import { shiftFormula } from './refs';
import { ERROR_NUM, ERROR_VALUE, arrayAt, calcError, isArray, isError, isRef, parseErrorCode, refToArray, type ArrayValue, type CalcContext, type CalcError, type FunctionSpec, type Scalar, type Value } from './values';

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
}

export interface RecalculateOptions {
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
  /** Every formula that fills more than its own cell. */
  spills: Set<FormulaRec>;
  /** The ranges of several cells that formulas read on this sheet, one node per range. */
  ranges: Map<string, RangeNode>;
  /** Formulas by column, built when circular references are looked for; undefined after the formulas changed. */
  columns?: Map<number, Column>;
  /** Keys of `columns` in ascending order. */
  columnKeys: number[];
  spillList: FormulaRec[];
  /** Row indexes in ascending order, built when a large range is first read in a recalculation. */
  rowKeys?: number[];
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

/** Functions the evaluator knows by name but cannot calculate. */
const UNSUPPORTED_FORMS = new Set(['LET', 'LAMBDA']);

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
  private batching = false;
  private batchOpen = false;

  constructor(
    private readonly workbook: Workbook,
    private readonly options: CalcEngineOptions = {},
  ) {
    this.functions = options.functions || FUNCTIONS;
    const engine = this;
    this.context = {
      sheet: 0,
      row: 0,
      col: 0,
      get date1904() {
        return !!workbook.date1904;
      },
      cellValue: (sheet, row, col) => engine.cellValue(sheet, row, col),
      forEachCell: (sheet, range, visit) => engine.forEachCell(sheet, range, visit),
      usedExtent: (sheet) => {
        const model = workbook.sheets[sheet];
        return model ? { rows: model.rowCount, cols: model.colCount } : { rows: 0, cols: 0 };
      },
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
   * Cells the engine cannot calculate are not listed once it has looked at them.
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
            // Not in the graph: it cannot be parsed, or its master is missing. Over budget: it was tried.
            if (rec ? rec.formula === cell.formula && rec.costly : !this.wantCells.some((want) => want.sheet === s && want.row === index && want.col === col)) return;
          }
          out.push({ sheet: s, row: index, col });
        });
      });
    }
    return out;
  }

  /**
   * Recalculates what depends on the given cells (all formulas when `changed` is omitted).
   * Writes results into the model by replacing cell objects (never mutating them) and
   * marks sheet.changedRows / sheet.modified for cells whose value changed.
   * Returns the cells whose value changed, per sheet.
   *
   * `changed` has to name every cell whose content or formula was changed since the last
   * call. Never throws: a formula that cannot be calculated gets an error value.
   *
   * The call stops when its budget (options.budget) is used up and sets `incomplete`. What
   * was calculated by then is written; the other formulas keep their previous values, never
   * a half-calculated one, and are taken up by the next call, whatever its `changed` is.
   * Nothing is written before the dependency graph is complete. A formula finished by one
   * call is not calculated again by the calls that go on with the rest.
   */
  recalculate(changed?: CellAddress[], options?: RecalculateOptions): { sheet: number; cells: { row: number; col: number }[] }[] {
    this.prepare();
    this.batchOpen = false;
    this.startBudget();
    for (const sh of this.sheets) {
      sh.copied.clear();
      sh.changed = [];
    }
    if (!Array.isArray(changed)) {
      this.wantAll = true;
      if (options && options.overwrite === true) this.wantOverwrite = true;
    } else {
      for (const cell of changed) {
        if (!cell || !this.sheets[cell.sheet] || !isIndex(cell.row, MAX_ROWS) || !isIndex(cell.col, MAX_COLS)) continue;
        this.wantCells.push({ sheet: cell.sheet, row: cell.row, col: cell.col });
      }
    }
    this.incomplete = true;
    if (!this.build()) return [];
    // The graph follows the cells that were changed; only then can circular references be looked for.
    for (const cell of this.wantCells) {
      const sh = this.sheets[cell.sheet];
      const before = sh.recs.get(keyOf(cell.row, cell.col));
      this.refresh(sh, cell.row, cell.col);
      const after = sh.recs.get(keyOf(cell.row, cell.col));
      this.ready.push({ sheet: cell.sheet, row: cell.row, col: cell.col, replaced: after !== undefined && after !== before });
    }
    this.wantCells = [];
    if (!this.components()) return [];

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
    this.calculateAll();
    this.flush();
    // Formulas that turned out to have been changed behind the engine's back are read again for next time.
    for (const rec of this.outdated) this.refresh(this.sheets[rec.sheet], rec.row, rec.col);
    this.outdated = [];
    const dirty = this.dirty;
    while (this.cursor < dirty.length && (dirty[this.cursor].state === DONE || dirty[this.cursor].epoch !== this.epoch)) this.cursor++;
    this.incomplete = this.first !== undefined || this.cursor < dirty.length;
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
    for (const sh of this.sheets) sh.rowKeys = undefined;
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
    // Which rows are hidden is something these formulas read: their results replace the values of the file.
    for (const rec of rowReaders) if (!rec.gone) this.mark(rec, true);
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
    for (let k = 0; k < this.dirty.length; k++) if (this.dirty[k].spill) this.materialize(this.dirty[k]);
  }

  /**
   * Call after rows or columns were inserted or deleted, sheets changed, or defined names
   * changed: forgets the dependency graph. The recalculation that follows builds it again,
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
      // Cut to what is in use, as a reference read as an array is; at least its first cell stays.
      const bottom = Math.min(range.bottom, Math.max((sh.model.rowCount || 0) - 1, range.top));
      const right = Math.min(range.right, Math.max((sh.model.colCount || 0) - 1, range.left));
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
    } else if (sh.spillCells.size > 0 || sh.spillRanges.length > 0) {
      const owner = this.spillOwner(sh, row, col);
      if (owner && owner.epoch === this.epoch) return this.spillValue(owner, row, col, cell);
    }
    return modelValue(cell, !!this.workbook.date1904);
  }

  private forEachCell(sheet: number, range: CellRange, visit: (value: Scalar, row: number, col: number) => void): void {
    const sh = this.sheets[sheet];
    if (!sh || !range) return;
    const rows = sh.model.rows;
    const top = Math.max(range.top, 0);
    const bottom = Math.min(range.bottom, MAX_ROWS - 1);
    const left = Math.max(range.left, 0);
    const right = Math.min(range.right, MAX_COLS - 1);
    if (!(top <= bottom && left <= right)) return;
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
    const range = rec.spill as CellRange;
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
    try {
      const value = evaluate(rec.node, context, this.functions, this.evalOptions);
      if (!rec.isArray) return rec.arrayMode ? this.firstOf(value) : toCellValue(value, context);
      const result = isRef(value) ? refToArray(value, context) : isArray(value) ? value : { rows: 1, cols: 1, data: [value] };
      rec.result = result;
      return result.data.length > 0 ? result.data[0] : ERROR_VALUE;
    } catch (error) {
      if (error instanceof EvaluationAbort) throw error;
      return ERROR_VALUE;
    }
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
    if (!rec.live) {
      const stored = this.storedOf(rec);
      if (stored !== undefined) {
        keep = sameValue(stored, value) && !(rec.spill && this.spillDiffers(rec)) ? KEEP_SAME : KEEP_STORED;
        value = stored;
      }
    }
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
    if (++this.finished === 1 && this.extended) {
      // The formula the budget was extended for is done: the next look at the budget ends the call.
      this.workLimit = this.softWork;
      this.timeLimit = this.softTime;
      this.nextCheck = 0;
    }
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
    for (let k = 0; k < done.length; k++) {
      this.store(done[k]);
      done[k].live = false;
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
    if (rec.keep === KEEP_NONE) this.put(sh, rec.row, rec.col, cell, rec.value);
    const range = rec.spill;
    if (!range) return;
    const result = rec.result || ZERO;
    const zero = rec.circular || rec.looped;
    for (let r = range.top; r <= range.bottom; r++) {
      for (let c = range.left; c <= range.right; c++) {
        if (r === rec.row && c === rec.col) continue;
        const row = sh.model.rows.get(r);
        const existing = row ? row.cells[c] : undefined;
        // A formula of its own is not what the array formula left there.
        if (existing && existing.formula) continue;
        // Where the result does not reach, arrayAt() answers #N/A, as Excel fills such cells.
        const value = zero ? 0 : arrayAt(result, r - range.top, c - range.left);
        this.put(sh, r, c, existing, value);
      }
    }
  }

  /**
   * Makes sure the cells an array formula fills exist, so that a range read while calculating
   * meets them. They get their values when the results are written.
   */
  private materialize(rec: FormulaRec): void {
    const sh = this.sheets[rec.sheet];
    const model = sh.model;
    const range = rec.spill as CellRange;
    for (let r = range.top; r <= range.bottom; r++) {
      let row = model.rows.get(r);
      for (let c = range.left; c <= range.right; c++) {
        if (row && row.cells[c]) continue;
        if (!row) {
          row = { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] };
          model.rows.set(r, row);
          sh.copied.add(r);
          sh.rowKeys = undefined;
        } else if (!sh.copied.has(r)) {
          row = { ...row, cells: row.cells.slice() };
          model.rows.set(r, row);
          sh.copied.add(r);
        }
        row.cells[c] = { kind: 'blank', value: null, style: 0 };
        if (r >= model.rowCount) model.rowCount = r + 1;
        if (c >= model.colCount) model.colCount = c + 1;
      }
    }
  }

  /** Puts a value into a cell by replacing the cell and its row, when it differs from what is there. */
  private put(sh: SheetState, row: number, col: number, cell: Cell | undefined, value: Scalar): void {
    let kind: Cell['kind'];
    let stored: number | string | boolean;
    if (value === null || value === undefined) (kind = 'number'), (stored = 0);
    else if (typeof value === 'number') (kind = 'number'), (stored = isFinite(value) ? value : 0);
    else if (typeof value === 'string') (kind = 'string'), (stored = value);
    else if (typeof value === 'boolean') (kind = 'boolean'), (stored = value);
    else if (isError(value)) (kind = 'error'), (stored = value.error);
    else (kind = 'error'), (stored = '#VALUE!');
    if (cell && cell.kind === kind && cell.value === stored) return;
    const next: Cell = cell ? { ...cell, kind, value: stored } : { kind, value: stored, style: 0 };
    // The text no longer comes from the shared string the cell was read from.
    delete next.runs;
    delete next.si;
    this.own.add(next);
    const model = sh.model;
    let target = model.rows.get(row);
    if (!target) {
      target = { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] };
      model.rows.set(row, target);
      sh.copied.add(row);
      sh.rowKeys = undefined;
    } else if (!sh.copied.has(row)) {
      target = { ...target, cells: target.cells.slice() };
      model.rows.set(row, target);
      sh.copied.add(row);
    }
    target.cells[col] = next;
    (model.changedRows || (model.changedRows = new Set())).add(row);
    model.modified = true;
    if (row >= model.rowCount) model.rowCount = row + 1;
    if (col >= model.colCount) model.colCount = col + 1;
    sh.changed.push({ row, col });
  }

  // -------------------------------------------------------------------------
  // The dependency graph

  /** Sheets and names: all that evaluating one formula needs. */
  private prepare(): void {
    if (this.prepared) return;
    this.prepared = true;
    const workbook = this.workbook;
    this.sheets = (workbook.sheets || []).map((model, index) => ({
      index,
      model,
      recs: new Map(),
      cellDeps: new Map(),
      rowBands: new Map(),
      colBands: new Map(),
      wide: { entries: [], done: 0, doneLive: 0 },
      spillCells: new Map(),
      spillRanges: [],
      spills: new Set(),
      ranges: new Map(),
      columnKeys: [],
      spillList: [],
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
    if (formula.type === 'dataTable') return;
    let text = formula.text;
    let dRow = 0;
    let dCol = 0;
    if (formula.type === 'shared' && text === '') {
      const master = formula.si === undefined ? undefined : (holder.masters || (holder.masters = sharedMasters(sh.model))).get(formula.si);
      // Without its master the cell cannot be calculated; it keeps its value.
      if (!master) return;
      text = master.text;
      dRow = row - master.row;
      dCol = col - master.col;
    }
    let node: Node;
    try {
      node = parseFormula(text);
    } catch (error) {
      this.unparsedCells.set(`${sh.index}:${row}:${col}`, { sheet: sh.index, row, col });
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
      isArray: formula.type === 'array',
      // Excel marks the cell of a dynamic array formula with cell metadata.
      arrayMode: formula.type === 'array' || cell.cm !== undefined,
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
    if (rec.isArray && typeof formula.ref === 'string') {
      const range = parseRange(formula.ref);
      const inside = range && row >= range.top && row <= range.bottom && col >= range.left && col <= range.right;
      if (range && inside) {
        const cells = (range.bottom - range.top + 1) * (range.right - range.left + 1);
        if (cells > 1 && cells <= MAX_SPILL_CELLS) {
          rec.spill = range;
          sh.spills.add(rec);
          if (cells <= MAX_LISTED_SPILL) {
            for (let r = range.top; r <= range.bottom; r++) for (let c = range.left; c <= range.right; c++) sh.spillCells.set(keyOf(r, c), rec);
          } else sh.spillRanges.push(rec);
        }
      }
    }
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

  /**
   * Whether everything a formula uses can be calculated here: no function the library lacks,
   * no LET or LAMBDA, no reference across sheets or into another workbook, no sheet or name
   * the workbook does not have (the name of a table is one: tables are not in the model), no
   * defined name that cannot be parsed or uses one of those, and no use of a function that
   * the evaluator only knows in part.
   */
  private supports(root: Node, sheet: number): boolean {
    const names = new Set<string>();
    const nodes: Node[] = [root];
    while (nodes.length > 0) {
      const node = nodes.pop() as Node;
      switch (node.type) {
        case 'ref':
          if (node.external || node.sheetEnd !== undefined) return false;
          if (node.sheet !== undefined && this.sheetIndex(node.sheet) === undefined) return false;
          break;
        case 'name': {
          this.metNames.push(node.name.toUpperCase());
          const home = node.sheet === undefined ? sheet : this.sheetIndex(node.sheet);
          if (home === undefined) return false;
          const key = `${home}!${node.name.toUpperCase()}`;
          if (names.has(key)) break;
          names.add(key);
          const target = this.nameTree(node.name, home);
          if (!target) return false;
          nodes.push(target);
          break;
        }
        case 'call': {
          const builtin = isBuiltinForm(node.name);
          const spec = builtin ? undefined : this.functions.get(node.name);
          if (UNSUPPORTED_FORMS.has(node.name) || (!builtin && !spec)) return false;
          if (!this.supportsCall(node.name, node.args)) return false;
          // A union of references is taken apart by AREAS and by functions of any number of arguments.
          const takesUnion = node.name === 'AREAS' || (spec !== undefined && spec.max === undefined);
          for (const arg of node.args) {
            let part = arg;
            while (takesUnion && part && part.type === 'binary' && part.op === ',') {
              nodes.push(part.right);
              part = part.left;
            }
            if (part) nodes.push(part);
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
      case 'INDEX': {
        // Only the first area of a reference.
        const area = args.length > 3 ? args[3] : undefined;
        return !area || (area.type === 'number' && area.value === 1);
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
          member.circular = member.selfRef;
          if (member !== rec) this.seeds.push(member);
        }
      }
    }
    if (rec.spill) {
      const range = rec.spill;
      const at = sh.spillRanges.indexOf(rec);
      if (at >= 0) sh.spillRanges.splice(at, 1);
      else {
        for (let r = range.top; r <= range.bottom; r++) {
          for (let c = range.left; c <= range.right; c++) if (sh.spillCells.get(keyOf(r, c)) === rec) sh.spillCells.delete(keyOf(r, c));
        }
      }
    }
    const drop = (bucket: Bucket | undefined) => {
      if (bucket) bucket.entries = bucket.entries.filter((entry) => entry.rec !== rec);
    };
    for (const precedent of rec.precedents) {
      const target = this.sheets[precedent.sheet];
      if (!target) continue;
      const range = precedent.range;
      if (range.top === range.bottom && range.left === range.right) {
        const key = keyOf(range.top, range.left);
        const list = target.cellDeps.get(key);
        if (list) {
          const kept = list.filter((other) => other !== rec);
          if (kept.length > 0) target.cellDeps.set(key, kept);
          else target.cellDeps.delete(key);
        }
        continue;
      }
      const firstBand = range.top >> BAND_SHIFT;
      const lastBand = range.bottom >> BAND_SHIFT;
      if (lastBand - firstBand < MAX_BANDS) for (let band = firstBand; band <= lastBand; band++) drop(target.rowBands.get(band));
      else if (range.right - range.left < MAX_BANDS) for (let c = range.left; c <= range.right; c++) drop(target.colBands.get(c));
      else drop(target.wide);
    }
    rec.precedents = [];
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
          (rec.names || (rec.names = [])).push(node.name.toUpperCase());
          const sheet = node.sheet === undefined ? rec.sheet : this.sheetIndex(node.sheet);
          if (sheet === undefined) break;
          const key = `${sheet}!${node.name.toUpperCase()}:${shapeOnly ? 1 : 0}`;
          if (names.has(key)) break;
          names.add(key);
          const target = this.resolveName(node.name, sheet);
          if (target) {
            nodes.push(target);
            own.push(false);
            shape.push(shapeOnly && target.type === 'ref');
            picked.push(pickedFrom && target.type === 'ref');
          }
          break;
        }
        case 'call': {
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
      run = { all: this.searchAll, seeds, next: 0, stamp: ++this.stamps, counter: 0, stack: [], path: [], a: [], b: [], c: [], steps: 0, limit: 64 * (total + seeds.length) + 4000000 };
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
      // Formulas elsewhere that fill cells of the range.
      const spills = sh.spillList;
      let si = run.c[frame];
      while (si < spills.length) {
        const rec = spills[si++];
        const spill = rec.spill as CellRange;
        if (spill.top <= range.bottom && spill.bottom >= range.top && spill.left <= range.right && spill.right >= range.left) {
          run.c[frame] = si;
          return rec;
        }
      }
      run.c[frame] = si;
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
      if (sh.spillCells.size > 0 || sh.spillRanges.length > 0) {
        const owner = this.spillOwner(sh, row, col);
        if (owner && owner !== inCell) {
          run.a[frame] = pi;
          return owner;
        }
      }
    }
    run.a[frame] = pi;
    return undefined;
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
        rec.circular = true;
      }
      this.groups.set(id, members);
    } else if (members.length === 1) {
      // Alone: circular only when it names its own cell.
      const rec = members[0];
      if (rec.comp) this.groups.delete(rec.comp);
      rec.comp = 0;
      rec.circular = rec.selfRef;
    }
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
    sh.spillList = Array.from(sh.spills);
  }

  // -------------------------------------------------------------------------
  // Marking what has to be calculated

  /** Marks a formula dirty for this recalculation. */
  private touch(rec: FormulaRec): void {
    if (rec.epoch === this.epoch) return;
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
