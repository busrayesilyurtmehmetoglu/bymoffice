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
// unknown function, a reference across sheets or into another workbook), and those that
// cannot be calculated within the budget of a recalculation.

import { isoToSerial } from '../format/cell-format';
import { MAX_COLS, MAX_ROWS, parseRange, type Cell, type CellFormula, type CellRange, type Row, type Sheet, type Workbook } from '../model/sheet/types';
import { sharedMasters, type SharedMasters } from '../ops/ops';
import { EvaluationAbort, evaluate, evaluationDepth, isBuiltinForm, refRange, toCellValue, type EvaluateOptions } from './eval';
import { FUNCTIONS } from './functions/index';
import { parseFormula, type Node } from './parse';
import { shiftFormula } from './refs';
import { ERROR_NUM, ERROR_VALUE, arrayAt, calcError, isArray, isError, isRef, parseErrorCode, refToArray, type ArrayValue, type CalcContext, type FunctionSpec, type Scalar, type Value } from './values';

/** How much one recalculate() call may cost before it stops with part of the formulas left for the next call. */
export interface CalcBudget {
  /** Milliseconds on the clock. 5000 when omitted; 0 or Infinity for no limit. */
  milliseconds?: number;
  /** Units of work: cells read, elements of arrays calculated, functions called. No limit when omitted. */
  work?: number;
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

export interface CellAddress {
  sheet: number;
  row: number;
  col: number;
}

// State of a formula within one recalculation.
const FRESH = 0;
const ACTIVE = 1;
const DONE = 2;

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
  /** Calculated in every recalculation: reads the clock, a random number or cells chosen while calculating. */
  volatile: boolean;
  isArray: boolean;
  /** Ranges are arrays everywhere in the formula: an array formula or a dynamic array formula. */
  arrayMode: boolean;
  /** One of its references names its own cell. */
  selfRef: boolean;
  /** How often a recalculation ran out of budget while calculating this formula. */
  aborts: number;
  /** Could not be calculated within a whole budget; left alone until its formula changes. */
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
  /** Part of a circular reference. */
  circular: boolean;
  /** The cell or its formula changed without the engine being told; skipped until the record is rebuilt. */
  stale: boolean;
  /** Index in the chain of formulas being calculated. */
  position: number;
}

interface Precedent {
  sheet: number;
  range: CellRange;
  /** Only asked where it is (ROW, COLUMNS, the base of OFFSET): its cells are not read. */
  shapeOnly: boolean;
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
  /** Keys of `recs` in ascending order, built when a circular reference is looked into. */
  recKeys?: number[];
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

/** Calculated every time, whatever changed. */
const VOLATILE = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'RANDARRAY', 'OFFSET', 'INDIRECT']);

/** Functions that only ask where a reference is; the cells of a reference given to them directly are not read. */
const SHAPE_ONLY = new Set(['ROW', 'COLUMN', 'ROWS', 'COLUMNS', 'AREAS', 'SHEET', 'SHEETS', 'ISREF', 'ISFORMULA', 'FORMULATEXT']);

/** Functions the evaluator knows by name but cannot calculate. */
const UNSUPPORTED_FORMS = new Set(['LET', 'LAMBDA']);

const DEFAULT_MILLISECONDS = 5000;
/** Units of work between two looks at the clock. */
const CLOCK_EVERY = 2048;

/** Thrown through the evaluator to start over with a shallower stack. */
const UNWIND = new EvaluationAbort();
/** Thrown when the budget of a recalculation is used up. */
const BUDGET = new EvaluationAbort();
/** Thrown when a circular reference is found among formulas that were already calculated: the pass starts over. */
const RESTART = new EvaluationAbort();
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

export class CalcEngine {
  /** Cells found in a circular reference by the last recalculation; they hold 0. */
  circular: CellAddress[] = [];
  /**
   * The last recalculate() call ran out of budget: some formulas were not calculated and
   * still hold their previous values. Call recalculate([]) (or with the next changes) to go on.
   */
  incomplete = false;

  private readonly functions: Map<string, FunctionSpec>;
  private readonly context: MutableContext;
  private readonly evalOptions: EvaluateOptions;
  private sheets: SheetState[] = [];
  private sheetByName = new Map<string, number>();
  private names = new Map<string, { global?: NameEntry; local?: Map<number, NameEntry> }>();
  private volatiles = new Set<FormulaRec>();
  private unparsedCells = new Map<string, CellAddress>();
  private unsupportedCells = new Map<string, CellAddress>();
  private costlyCells = new Map<string, CellAddress>();
  /** Formulas a recalculation that ran out of budget left over, in the order to take them up again. */
  private carry: FormulaRec[] = [];
  /** What functions keep for the rest of a recalculation (see CalcContext.memo). */
  private memos = new Map<string, unknown>();
  private work = 0;
  private workLimit = Infinity;
  private timeLimit = Infinity;
  private started = 0;
  /** Amount of work at which the budget is looked at next. */
  private nextCheck = Infinity;
  private prepared = false;
  private built = false;
  private epoch = 0;
  private nowValue: number | undefined;
  private dirty: FormulaRec[] = [];
  private stale: FormulaRec[] = [];
  /** Formulas being calculated, each one needed by the one before it. */
  private chain: FormulaRec[] = [];
  /** Length of the chain when the formula now on top of the JavaScript stack was started. */
  private floor = 0;
  private pending: FormulaRec | undefined;

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
   * or a name the workbook does not have, a defined name that uses one of those. They keep
   * the values the file has for them, and formulas that read them read those values. A
   * formula without a stored value (one just typed in) is not listed: it is calculated and
   * shows the error (#NAME?, #REF!).
   */
  get unsupported(): CellAddress[] {
    return Array.from(this.unsupportedCells.values());
  }

  /**
   * Formula cells that could not be calculated within a whole budget, twice. They keep the
   * values they had and are not tried again until their formula changes or invalidate() is called.
   */
  get overBudget(): CellAddress[] {
    return Array.from(this.costlyCells.values());
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
   */
  recalculate(changed?: CellAddress[]): { sheet: number; cells: { row: number; col: number }[] }[] {
    this.build();
    this.epoch++;
    this.nowValue = undefined;
    this.circular = [];
    this.dirty = [];
    this.stale = [];
    this.memos.clear();
    this.startBudget();
    for (const sh of this.sheets) {
      sh.rowKeys = undefined;
      sh.copied.clear();
      sh.changed = [];
    }
    const carried = this.carry;
    this.carry = [];
    this.incomplete = false;

    if (!Array.isArray(changed)) {
      // What the last call left over comes first, so that every formula gets its turn at a fresh budget.
      for (const rec of carried) if (!rec.stale) this.touch(rec);
      for (const sh of this.sheets) sh.recs.forEach((rec) => this.touch(rec));
    } else {
      const first = this.dirty.length;
      for (const rec of carried) if (!rec.stale) this.touch(rec);
      for (const cell of changed) {
        if (!cell) continue;
        const sh = this.sheets[cell.sheet];
        if (!sh || !isIndex(cell.row, MAX_ROWS) || !isIndex(cell.col, MAX_COLS)) continue;
        this.refresh(sh, cell.row, cell.col);
        const rec = sh.recs.get(keyOf(cell.row, cell.col));
        if (rec) this.touch(rec);
        else this.collect(sh, cell.row, cell.col, cell.row, cell.col);
      }
      this.volatiles.forEach((rec) => this.touch(rec));
      // Everything that reads a dirty formula is dirty as well; `dirty` grows while it is walked.
      for (let k = first; k < this.dirty.length; k++) {
        const rec = this.dirty[k];
        const sh = this.sheets[rec.sheet];
        if (rec.spill) this.collect(sh, rec.spill.top, rec.spill.left, rec.spill.bottom, rec.spill.right);
        else this.collect(sh, rec.row, rec.col, rec.row, rec.col);
      }
    }

    for (let k = 0; k < this.dirty.length; k++) if (this.dirty[k].spill) this.materialize(this.dirty[k]);
    this.calculateAll();
    for (let k = 0; k < this.dirty.length; k++) this.store(this.dirty[k]);
    // Formulas that turned out to have been changed behind the engine's back are read again for next time.
    for (const rec of this.stale) this.refresh(this.sheets[rec.sheet], rec.row, rec.col);
    this.memos.clear();
    this.incomplete = this.carry.length > 0;

    const out: { sheet: number; cells: { row: number; col: number }[] }[] = [];
    for (const sh of this.sheets) if (sh.changed.length > 0) out.push({ sheet: sh.index, cells: sh.changed });
    return out;
  }

  /** Call after rows or columns were inserted or deleted, sheets changed, or defined names changed: forgets the dependency graph. */
  invalidate(): void {
    this.prepared = false;
    this.built = false;
    this.sheets = [];
    this.sheetByName = new Map();
    this.names = new Map();
    this.volatiles = new Set();
    this.unparsedCells = new Map();
    this.unsupportedCells = new Map();
    this.costlyCells = new Map();
    // Whatever was left over is calculated with everything else: the caller recalculates all after this.
    this.carry = [];
    this.incomplete = false;
    this.memos.clear();
    this.dirty = [];
    this.stale = [];
    this.chain = [];
  }

  /**
   * The value a formula text would have in a cell, without storing anything (for a formula bar preview).
   * Other formulas are not recalculated: they answer with the values the model holds.
   * Throws FormulaParseError when the text is not a formula, so that the caller can show
   * where it goes wrong; a formula that parses never throws.
   */
  evaluateText(formula: string, at: CellAddress): Scalar {
    const node = parseFormula(formula);
    this.prepare();
    // A new epoch: no formula is dirty, every cell answers from the model.
    this.epoch++;
    this.nowValue = undefined;
    this.memos.clear();
    this.startBudget();
    for (const sh of this.sheets) sh.rowKeys = undefined;
    const context = this.context;
    context.sheet = at && isIndex(at.sheet, this.sheets.length) ? at.sheet : 0;
    context.row = at && isIndex(at.row, MAX_ROWS) ? at.row : 0;
    context.col = at && isIndex(at.col, MAX_COLS) ? at.col : 0;
    this.evalOptions.rowShift = 0;
    this.evalOptions.colShift = 0;
    this.evalOptions.arrayFormula = false;
    let value: Scalar;
    try {
      value = toCellValue(evaluate(node, context, this.functions, this.evalOptions), context);
    } catch (error) {
      // Out of budget, among other things.
      value = ERROR_VALUE;
    }
    this.memos.clear();
    if (typeof value === 'number' && !isFinite(value)) return ERROR_NUM;
    return value === null || value === undefined ? 0 : value;
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

  private startBudget(): void {
    const budget = this.options.budget;
    const milliseconds = budget && budget.milliseconds !== undefined ? Number(budget.milliseconds) : DEFAULT_MILLISECONDS;
    const work = budget && budget.work !== undefined ? Number(budget.work) : Infinity;
    this.timeLimit = milliseconds > 0 ? milliseconds : Infinity;
    this.workLimit = work >= 0 ? work : Infinity;
    this.work = 0;
    this.started = this.timeLimit < Infinity ? this.clock() : 0;
    this.nextCheck = Math.min(CLOCK_EVERY, this.workLimit + 1);
  }

  private clock(): number {
    const value = this.options.clock ? Number(this.options.clock()) : Date.now();
    return value === value ? value : 0;
  }

  /** Counts work; gives the calculation up when the budget is used up. */
  private spend(units: number): void {
    this.work += units;
    if (this.work < this.nextCheck) return;
    if (this.work > this.workLimit) throw BUDGET;
    this.nextCheck = Math.min(this.work + CLOCK_EVERY, this.workLimit + 1);
    if (this.timeLimit < Infinity && this.clock() - this.started > this.timeLimit) throw BUDGET;
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

  private cellValue(sheet: number, row: number, col: number): Scalar {
    const sh = this.sheets[sheet];
    if (!sh) return null;
    if (++this.work >= this.nextCheck) this.spend(0);
    const modelRow = sh.model.rows.get(row);
    const cell = modelRow ? modelRow.cells[col] : undefined;
    if (!cell) {
      // A cell an array formula fills may not exist yet.
      if (sh.spillCells.size > 0 || sh.spillRanges.length > 0) {
        const owner = this.spillOwner(sh, row, col);
        if (owner && owner.epoch === this.epoch) return this.spillValue(owner, row, col);
      }
      return null;
    }
    return this.valueOf(sh, cell, row, col);
  }

  private valueOf(sh: SheetState, cell: Cell, row: number, col: number): Scalar {
    if (cell.formula !== undefined) {
      if (sh.rowEpoch[row] === this.epoch) {
        const rec = sh.recs.get(keyOf(row, col));
        if (rec !== undefined && rec.epoch === this.epoch && rec.formula === cell.formula) return this.need(rec);
      }
    } else if (sh.spillCells.size > 0 || sh.spillRanges.length > 0) {
      const owner = this.spillOwner(sh, row, col);
      if (owner && owner.epoch === this.epoch) return this.spillValue(owner, row, col);
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

  private spillValue(rec: FormulaRec, row: number, col: number): Scalar {
    this.need(rec);
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
  // Circular references. A formula that is asked for while it is being calculated is part
  // of a cycle. Which formulas were on the stack at that moment depends on where the
  // calculation started, so the cycle is not taken from the stack alone: every formula that
  // the references tie to the one found (each reads the other, directly or through others)
  // is marked, and all of them get 0, in whatever order the cells are calculated. Formulas
  // that only read a cycle are calculated normally.

  /** The value of a dirty formula, calculating it first if needed. */
  private need(rec: FormulaRec): Scalar {
    if (rec.state === DONE) return rec.value;
    const chain = this.chain;
    if (rec.circular) {
      if (rec.state !== ACTIVE) this.settle(rec);
      return 0;
    }
    if (rec.state === ACTIVE) {
      this.markCircular(rec);
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

  /** Calculates every dirty formula, until all are done or the budget is used up. */
  private calculateAll(): void {
    const dirty = this.dirty;
    for (;;) {
      try {
        for (let k = 0; k < dirty.length; k++) this.calculate(dirty[k]);
        return;
      } catch (error) {
        if (error === RESTART) {
          // Formulas calculated before the cycle was known may have read what is now 0: all over again.
          this.chain.length = 0;
          this.circular = [];
          this.memos.clear();
          for (let k = 0; k < dirty.length; k++) {
            dirty[k].state = FRESH;
            dirty[k].result = undefined;
          }
          continue;
        }
        this.abandon();
        return;
      }
    }
  }

  /** Out of budget: what is not done is left for the next recalculation. */
  private abandon(): void {
    const chain = this.chain;
    const blamed = chain.length > 0 ? chain[chain.length - 1] : undefined;
    chain.length = 0;
    const dirty = this.dirty;
    let done = 0;
    for (let k = 0; k < dirty.length; k++) if (dirty[k].state === DONE) done++;
    const carry: FormulaRec[] = [];
    if (blamed && blamed.state !== DONE) {
      blamed.aborts++;
      if (done === 0 && blamed.aborts >= 2) {
        // It had a whole budget to itself, twice. Trying again at every change would only freeze the caller.
        blamed.costly = true;
        blamed.epoch = 0;
        blamed.state = FRESH;
        this.costlyCells.set(`${blamed.sheet}:${blamed.row}:${blamed.col}`, { sheet: blamed.sheet, row: blamed.row, col: blamed.col });
      } else carry.push(blamed);
    }
    for (let k = 0; k < dirty.length; k++) {
      const rec = dirty[k];
      if (rec.state === DONE || rec.costly) continue;
      rec.state = FRESH;
      rec.result = undefined;
      if (rec !== blamed) carry.push(rec);
    }
    this.carry = carry;
  }

  private calculate(root: FormulaRec): void {
    if (root.state === DONE) return;
    const chain = this.chain;
    // Left over from a calculation that was cut off by an exception.
    for (const rec of chain) if (rec.state === ACTIVE) rec.state = FRESH;
    chain.length = 0;
    if (root.circular) {
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
        if (error === BUDGET || error === RESTART) throw error;
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
      if (!rec.stale) {
        rec.stale = true;
        this.stale.push(rec);
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
    if (rec.circular) {
      value = 0;
      rec.result = undefined;
      this.circular.push({ sheet: rec.sheet, row: rec.row, col: rec.col });
    }
    // An empty result shows as 0, as a reference to an empty cell does.
    if (value === null || value === undefined) value = 0;
    else if (typeof value === 'number' && !isFinite(value)) value = ERROR_NUM;
    rec.value = value;
    rec.state = DONE;
  }

  /** A formula known to be circular before it is calculated: it is 0. */
  private settle(rec: FormulaRec): void {
    rec.value = 0;
    rec.result = undefined;
    rec.state = DONE;
    this.circular.push({ sheet: rec.sheet, row: rec.row, col: rec.col });
  }

  /** `rec` was asked for while it was being calculated: marks every formula of its cycles. */
  private markCircular(rec: FormulaRec): void {
    const chain = this.chain;
    const members = this.component(rec);
    // What the stack shows is a cycle too, also when it runs through references made while calculating (OFFSET, INDIRECT).
    for (let k = rec.position; k < chain.length; k++) members.push(chain[k]);
    let restart = false;
    for (const member of members) {
      if (member.circular) continue;
      member.circular = true;
      if (member.state === DONE) restart = true;
    }
    if (restart) throw RESTART;
  }

  /**
   * The formulas tied to `start` by references in both directions: those it reads, directly
   * or through other formulas, that also read it. Worked out from the references the formulas
   * name, whether or not they are reached when calculating. Only dirty formulas matter: what
   * reads a dirty formula is dirty itself.
   */
  private component(start: FormulaRec): FormulaRec[] {
    interface Group {
      /** The formulas inside one precedent range. */
      members: FormulaRec[];
      /** The formulas that read that range. */
      users: FormulaRec[];
      walked: boolean;
    }
    const groups = new Map<string, Group>();
    const reached = new Set<FormulaRec>([start]);
    const stack: FormulaRec[] = [start];
    while (stack.length > 0) {
      const rec = stack.pop() as FormulaRec;
      for (const precedent of rec.precedents) {
        if (precedent.shapeOnly) continue;
        const range = precedent.range;
        const key = `${precedent.sheet}:${range.top}:${range.left}:${range.bottom}:${range.right}`;
        let group = groups.get(key);
        if (!group) {
          group = { members: this.formulasIn(precedent.sheet, range), users: [], walked: false };
          groups.set(key, group);
          for (const member of group.members) {
            if (reached.has(member)) continue;
            reached.add(member);
            stack.push(member);
          }
        }
        group.users.push(rec);
      }
    }
    // Back from `start`: whoever reads a range it lies in, and whoever reads those.
    const inside = new Map<FormulaRec, Group[]>();
    groups.forEach((group) => {
      for (const member of group.members) {
        const list = inside.get(member);
        if (list) list.push(group);
        else inside.set(member, [group]);
      }
    });
    const out: FormulaRec[] = [start];
    const seen = new Set<FormulaRec>(out);
    stack.push(start);
    while (stack.length > 0) {
      const list = inside.get(stack.pop() as FormulaRec);
      if (!list) continue;
      for (const group of list) {
        if (group.walked) continue;
        group.walked = true;
        for (const user of group.users) {
          if (seen.has(user)) continue;
          seen.add(user);
          out.push(user);
          stack.push(user);
        }
      }
    }
    return out;
  }

  /** The dirty formulas whose cell, or one of the cells they fill, lies in a range. */
  private formulasIn(sheet: number, range: CellRange): FormulaRec[] {
    const sh = this.sheets[sheet];
    const out: FormulaRec[] = [];
    if (!sh) return out;
    const epoch = this.epoch;
    if (range.top === range.bottom && range.left === range.right) {
      const rec = sh.recs.get(keyOf(range.top, range.left));
      if (rec && rec.epoch === epoch) out.push(rec);
    } else {
      let keys = sh.recKeys;
      if (!keys) {
        keys = Array.from(sh.recs.keys()).sort((a, b) => a - b);
        sh.recKeys = keys;
      }
      const last = keyOf(range.bottom, range.right);
      let k = lowerBound(keys, keyOf(range.top, range.left));
      const from = k;
      for (; k < keys.length && keys[k] <= last; k++) {
        const col = keys[k] % MAX_COLS;
        if (col < range.left || col > range.right) continue;
        const rec = sh.recs.get(keys[k]);
        if (rec && rec.epoch === epoch) out.push(rec);
      }
      this.spend(k - from);
    }
    if (sh.spills.size > 0) {
      this.spend(sh.spills.size);
      sh.spills.forEach((rec) => {
        const spill = rec.spill as CellRange;
        if (rec.epoch !== epoch || spill.top > range.bottom || spill.bottom < range.top || spill.left > range.right || spill.right < range.left) return;
        if (out.indexOf(rec) < 0) out.push(rec);
      });
    }
    return out;
  }

  // -------------------------------------------------------------------------
  // Writing results

  private store(rec: FormulaRec): void {
    // What a recalculation that ran out of budget did not reach keeps its value.
    if (rec.stale || rec.state !== DONE) return;
    const sh = this.sheets[rec.sheet];
    const modelRow = sh.model.rows.get(rec.row);
    const cell = modelRow ? modelRow.cells[rec.col] : undefined;
    if (!cell || cell.formula !== rec.formula) return;
    this.put(sh, rec.row, rec.col, cell, rec.value);
    const range = rec.spill;
    if (!range) return;
    const result = rec.result || ZERO;
    for (let r = range.top; r <= range.bottom; r++) {
      for (let c = range.left; c <= range.right; c++) {
        if (r === rec.row && c === rec.col) continue;
        const row = sh.model.rows.get(r);
        const existing = row ? row.cells[c] : undefined;
        // A formula of its own is not what the array formula left there.
        if (existing && existing.formula) continue;
        // Where the result does not reach, arrayAt() answers #N/A, as Excel fills such cells.
        const value = rec.circular ? 0 : arrayAt(result, r - range.top, c - range.left);
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
    const model = sh.model;
    let target = model.rows.get(row);
    if (!target) {
      target = { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] };
      model.rows.set(row, target);
      sh.copied.add(row);
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
      wide: { entries: [], done: 0 },
      spillCells: new Map(),
      spillRanges: [],
      spills: new Set(),
      rowEpoch: new Uint32Array(Math.min(Math.max(model.rowCount || 0, 0), MAX_ROWS) + 16),
      copied: new Set(),
      changed: [],
    }));
    this.sheetByName = new Map();
    this.sheets.forEach((sh) => {
      const key = foldName(sh.model.name);
      if (!this.sheetByName.has(key)) this.sheetByName.set(key, sh.index);
    });
    this.names = new Map();
    for (const name of workbook.definedNames || []) {
      if (!name || typeof name.name !== 'string' || typeof name.value !== 'string') continue;
      const key = name.name.toUpperCase();
      let found = this.names.get(key);
      if (!found) this.names.set(key, (found = {}));
      const entry: NameEntry = { text: name.value };
      if (name.localSheet === undefined) {
        if (!found.global) found.global = entry;
      } else {
        if (!found.local) found.local = new Map();
        if (!found.local.has(name.localSheet)) found.local.set(name.localSheet, entry);
      }
    }
  }

  private build(): void {
    this.prepare();
    if (this.built) return;
    this.built = true;
    for (const sh of this.sheets) {
      const masters = sharedMasters(sh.model);
      sh.model.rows.forEach((row, index) => {
        if (!isIndex(index, MAX_ROWS) || !row || !row.cells) return;
        row.cells.forEach((cell, col) => {
          if (cell && cell.formula && col < MAX_COLS) this.addFormula(sh, index, col, cell, masters);
        });
      });
    }
  }

  /** Reads the formula of a cell into the graph. */
  private addFormula(sh: SheetState, row: number, col: number, cell: Cell, masters: SharedMasters | undefined): void {
    const formula = cell.formula as CellFormula;
    // The inputs of a data table are substituted by Excel itself; its cells keep their values.
    if (formula.type === 'dataTable') return;
    let text = formula.text;
    let dRow = 0;
    let dCol = 0;
    if (formula.type === 'shared' && text === '') {
      const master = formula.si === undefined ? undefined : (masters || sharedMasters(sh.model)).get(formula.si);
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
    // (a formula just typed in) it is calculated anyway, so that the cell shows the error.
    const hasValue = cell.kind !== 'blank' && cell.value !== null && cell.value !== undefined;
    if (hasValue && !this.supports(node, sh.index)) {
      this.unsupportedCells.set(`${sh.index}:${row}:${col}`, { sheet: sh.index, row, col });
      return;
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
      isArray: formula.type === 'array',
      // Excel marks the cell of a dynamic array formula with cell metadata.
      arrayMode: formula.type === 'array' || cell.cm !== undefined,
      selfRef: false,
      aborts: 0,
      costly: false,
      precedents: [],
      epoch: 0,
      state: FRESH,
      value: null,
      circular: false,
      stale: false,
      position: 0,
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
    if (rec.volatile) this.volatiles.add(rec);
    sh.recs.set(keyOf(row, col), rec);
    sh.recKeys = undefined;
  }

  /**
   * Whether everything a formula uses can be calculated here: no function the library lacks,
   * no LET or LAMBDA, no reference across sheets or into another workbook, no sheet or name
   * the workbook does not have (the name of a table is one: tables are not in the model), no
   * defined name that cannot be parsed or uses one of those.
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
        case 'call':
          if (UNSUPPORTED_FORMS.has(node.name) || (!isBuiltinForm(node.name) && !this.functions.has(node.name))) return false;
          for (const arg of node.args) if (arg) nodes.push(arg);
          break;
        case 'unary':
          nodes.push(node.operand);
          break;
        case 'binary':
          nodes.push(node.left, node.right);
          break;
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
    this.unparsedCells.delete(`${sh.index}:${row}:${col}`);
    this.unsupportedCells.delete(`${sh.index}:${row}:${col}`);
    this.costlyCells.delete(`${sh.index}:${row}:${col}`);
    if (rec) this.removeFormula(sh, rec);
    if (cell && formula) this.addFormula(sh, row, col, cell, undefined);
  }

  private removeFormula(sh: SheetState, rec: FormulaRec): void {
    sh.recs.delete(keyOf(rec.row, rec.col));
    sh.recKeys = undefined;
    sh.spills.delete(rec);
    this.volatiles.delete(rec);
    // Whoever still holds the record (the list of dirty formulas) must not write it.
    rec.stale = true;
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
          const asksWhere = SHAPE_ONLY.has(node.name);
          for (let k = 0; k < node.args.length; k++) {
            const arg = node.args[k];
            if (!arg) continue;
            const reference = arg.type === 'ref' || arg.type === 'name';
            nodes.push(arg);
            own.push(shifted);
            shape.push((asksWhere || (k === 0 && node.name === 'OFFSET')) && reference);
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
    rec.precedents.push({ sheet, range, shapeOnly });
    if (range.top === range.bottom && range.left === range.right) {
      const key = keyOf(range.top, range.left);
      const list = target.cellDeps.get(key);
      if (list) list.push(rec);
      else target.cellDeps.set(key, [rec]);
      return;
    }
    const entry: Entry = { top: range.top, left: range.left, bottom: range.bottom, right: range.right, rec };
    const add = (map: Map<number, Bucket>, key: number) => {
      const bucket = map.get(key);
      if (bucket) {
        bucket.entries.push(entry);
        bucket.done = 0;
      } else map.set(key, { entries: [entry], done: 0 });
    };
    const firstBand = range.top >> BAND_SHIFT;
    const lastBand = range.bottom >> BAND_SHIFT;
    if (lastBand - firstBand < MAX_BANDS) for (let band = firstBand; band <= lastBand; band++) add(target.rowBands, band);
    else if (range.right - range.left < MAX_BANDS) for (let c = range.left; c <= range.right; c++) add(target.colBands, c);
    else {
      target.wide.entries.push(entry);
      target.wide.done = 0;
    }
  }

  /** Marks a formula dirty for this recalculation. */
  private touch(rec: FormulaRec): void {
    if (rec.epoch === this.epoch || rec.costly) return;
    rec.epoch = this.epoch;
    rec.state = FRESH;
    rec.circular = rec.selfRef;
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

  /** Marks every formula that reads a cell of the range. */
  private collect(sh: SheetState, top: number, left: number, bottom: number, right: number): void {
    const deps = sh.cellDeps;
    if (top === bottom && left === right) {
      const list = deps.get(keyOf(top, left));
      if (list) for (let k = 0; k < list.length; k++) this.touch(list[k]);
    } else if (deps.size > 0) {
      if ((bottom - top + 1) * (right - left + 1) <= deps.size) {
        for (let r = top; r <= bottom; r++) {
          for (let c = left; c <= right; c++) {
            const list = deps.get(keyOf(r, c));
            if (list) for (let k = 0; k < list.length; k++) this.touch(list[k]);
          }
        }
      } else {
        deps.forEach((list, key) => {
          const r = Math.floor(key / MAX_COLS);
          const c = key - r * MAX_COLS;
          if (r >= top && r <= bottom && c >= left && c <= right) for (let k = 0; k < list.length; k++) this.touch(list[k]);
        });
      }
    }
    if (sh.rowBands.size > 0) {
      const first = top >> BAND_SHIFT;
      const last = bottom >> BAND_SHIFT;
      if (last - first < sh.rowBands.size) {
        for (let band = first; band <= last; band++) this.scan(sh.rowBands.get(band), top, left, bottom, right);
      } else {
        sh.rowBands.forEach((bucket, band) => {
          if (band >= first && band <= last) this.scan(bucket, top, left, bottom, right);
        });
      }
    }
    if (sh.colBands.size > 0) {
      if (right - left < sh.colBands.size) {
        for (let c = left; c <= right; c++) this.scan(sh.colBands.get(c), top, left, bottom, right);
      } else {
        sh.colBands.forEach((bucket, c) => {
          if (c >= left && c <= right) this.scan(bucket, top, left, bottom, right);
        });
      }
    }
    this.scan(sh.wide, top, left, bottom, right);
  }

  private scan(bucket: Bucket | undefined, top: number, left: number, bottom: number, right: number): void {
    if (!bucket || bucket.done === this.epoch) return;
    const entries = bucket.entries;
    let clean = 0;
    for (let k = 0; k < entries.length; k++) {
      const entry = entries[k];
      if (entry.rec.epoch === this.epoch) continue;
      if (entry.top <= bottom && entry.bottom >= top && entry.left <= right && entry.right >= left) this.touch(entry.rec);
      else clean++;
    }
    // Nothing left to find here until the next recalculation.
    if (clean === 0) bucket.done = this.epoch;
  }
}
