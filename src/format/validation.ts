// Data validation: what may be typed into a cell, the message a cell shows
// when it is selected, and the list a cell offers its values in.
//
// The rules are read from the file and never changed here; the writer copies
// their source. Rows and columns inserted or deleted since the sheet was read
// are applied to the ranges and formulas when the rules are looked up, the way
// conditional formats do it (see ConditionalFormats).
//
// Formulas are calculated by the engine the caller provides. A rule that needs
// one and cannot get it (no engine, a formula the engine cannot read or does
// not know how to calculate) is not applied: the value is let through rather
// than refused for no reason. A formula that is calculated and gives an error
// is another matter: like Excel, a rule takes that for a "no".

import type { CalcEngine, CellAddress } from '../calc/engine';
import { parseFormula, type Node } from '../calc/parse';
import { adjustAnchoredFormula, adjustRange, renameSheetsInFormula, type StructuralEdit } from '../calc/refs';
import { isError, type Scalar } from '../calc/values';
import { MAX_COLS, MAX_ROWS, fileSheetName, getCell, sheetNameMap, type Cell, type CellRange, type DataValidation, type Row, type Sheet, type Workbook } from '../model/sheet/types';
import { resolveCellStyle } from '../model/sheet/color';
import { applyOp, type Op } from '../ops/ops';
import { foldCase } from '../ops/sort';
import { builtinFormatId, formatCell, isoToSerial } from './cell-format';
import type { FormatLocale } from './locale';
import { formatValue } from './number-format';

/** Most values the list of a cell offers. */
export const MAX_LIST_ITEMS = 10000;

/** Most cells of a list source that are looked at to find a value. */
const MAX_SOURCE_CELLS = 2000000;

/** Defined names that only point at other names are followed this far. */
const MAX_NAME_DEPTH = 8;

export interface ValidationOptions {
  /** The formula engine, made on first use; without one, rules that need formulas are not applied. */
  engine?: () => CalcEngine | undefined;
  locale?: FormatLocale;
}

export interface ValidationVerdict {
  ok: boolean;
  /** How strict the rule is: "stop" refuses the value, the others only tell. */
  style?: 'stop' | 'warning' | 'information';
  /** Title and text of the rule's own message, as the file states them; either may be missing. */
  title?: string;
  message?: string;
}

export interface ValidationList {
  /** The values, as the cells of the source show them. */
  items: string[];
  /** The source has more values than the list holds. */
  truncated: boolean;
}

/** A value refused by a rule of style "stop". `display` is what to tell the user: the rule's own words. */
export class ValidationError extends RangeError {
  readonly display: string;

  constructor(display: string) {
    super('The value does not match the data validation of the cell');
    this.name = 'ValidationError';
    this.display = display;
  }
}

interface Entry {
  rule: DataValidation;
  ranges: CellRange[];
  anchorRow: number;
  anchorCol: number;
  formula1: string | undefined;
  formula2: string | undefined;
}

/** Where the values of a list come from. */
type Source =
  | { kind: 'constant'; items: string[] }
  /** `named`: the rule states a defined name, not the cells themselves. */
  | { kind: 'cells'; sheet: Sheet; range: CellRange; named?: boolean }
  | { kind: 'formula'; text: string };

type Plain = number | string | boolean;

const OK: ValidationVerdict = { ok: true };

/** What a limit is worth when its formula gives an error: no value is within it. */
const NO_LIMIT = 'error';

/**
 * Whether a formula uses what the engine cannot calculate, however it is written: a
 * reference into another workbook, or across sheets (3D). The engine answers those with
 * #REF!, which would read as a formula that was calculated and failed.
 */
const beyondEngine = new WeakMap<Node, boolean>();

function isBeyondEngine(root: Node): boolean {
  const known = beyondEngine.get(root);
  if (known !== undefined) return known;
  let found = false;
  // The parser bounds the size of a tree; the walk keeps its own stack all the same.
  const stack: Node[] = [root];
  while (stack.length > 0 && !found) {
    const node = stack.pop() as Node;
    switch (node.type) {
      case 'ref':
        found = node.external === true || node.sheetEnd !== undefined;
        break;
      case 'call':
        for (const arg of node.args) if (arg) stack.push(arg);
        break;
      case 'unary':
        stack.push(node.operand);
        break;
      case 'binary':
        stack.push(node.left, node.right);
        break;
      default:
      // Constants and names hold nothing more.
    }
  }
  beyondEngine.set(root, found);
  return found;
}

/** A number as files write it; undefined for any other text. */
function plainNumber(text: string): number | undefined {
  const trimmed = text.trim();
  if (trimmed.length === 0 || trimmed.length > 32 || !/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(trimmed)) return undefined;
  const value = Number(trimmed);
  return isFinite(value) ? value : undefined;
}

function inRanges(ranges: CellRange[], row: number, col: number): boolean {
  for (let i = 0; i < ranges.length; i++) {
    const range = ranges[i];
    if (row >= range.top && row <= range.bottom && col >= range.left && col <= range.right) return true;
  }
  return false;
}

/** The value a cell holds, as rules compare it; undefined for an empty cell and for an error. */
function plainValue(cell: Cell | undefined, date1904: boolean): Plain | undefined {
  if (!cell || cell.value === null || cell.value === undefined || cell.kind === 'blank' || cell.kind === 'error') return undefined;
  if (cell.kind === 'date' && typeof cell.value === 'string') {
    const serial = isoToSerial(cell.value, date1904);
    return serial === undefined ? cell.value : serial;
  }
  return cell.value;
}

/** The items of a list written into the rule: "a,b,c" in quotes. */
function constantItems(formula: string): string[] | undefined {
  const text = formula.trim();
  if (text.length < 2 || text.charAt(0) !== '"' || text.charAt(text.length - 1) !== '"') return undefined;
  const inner = text.slice(1, -1);
  // One string and nothing else: a quote inside it is doubled.
  if (/(^|[^"])"([^"]|$)/.test(inner.replace(/""/g, ''))) return undefined;
  // In the file the items are always separated by commas, whatever the locale of its author.
  return inner
    .replace(/""/g, '"')
    .split(',')
    .map((item) => item.trim())
    .filter((item) => item !== '');
}

function compare(value: number, operator: string, a: number, b: number | undefined): boolean {
  switch (operator) {
    case 'between':
      return b === undefined ? value >= a : value >= Math.min(a, b) && value <= Math.max(a, b);
    case 'notBetween':
      return b === undefined ? value < a : value < Math.min(a, b) || value > Math.max(a, b);
    case 'equal':
      return value === a;
    case 'notEqual':
      return value !== a;
    case 'greaterThan':
      return value > a;
    case 'lessThan':
      return value < a;
    case 'greaterThanOrEqual':
      return value >= a;
    case 'lessThanOrEqual':
      return value <= a;
    default:
      // An operator this engine does not know: nothing is refused for it.
      return true;
  }
}

/**
 * Stands for the name in the file of a sheet that is not in the file: no reference of the
 * file names it (a sheet name holds no control character), references without a sheet name
 * still follow the sheet's own edits.
 */
const NOT_IN_FILE = '\u0000';

/** The data validations of one sheet, as they apply now. */
export class DataValidations {
  private readonly workbook: Workbook;
  private readonly sheet: Sheet;
  private readonly options: ValidationOptions;
  private entries: Entry[] | undefined;
  private rulesSeen: DataValidation[] | undefined;
  /** The structuralEdits lists of all sheets the entries were built for. */
  private editsSeen: (Sheet['structuralEdits'] | undefined)[] = [];
  /** The sheets, their names and the deleted sheets the entries were built for. */
  private sheetsSeen: Sheet[] = [];
  private namesSeen: string[] = [];
  private removedSeen: string[] | undefined;

  constructor(workbook: Workbook, sheet: Sheet, options: ValidationOptions = {}) {
    this.workbook = workbook;
    this.sheet = sheet;
    this.options = options || {};
  }

  /** The sheet has at least one rule. */
  get active(): boolean {
    const rules = this.sheet.dataValidations;
    return Array.isArray(rules) && rules.length > 0;
  }

  /** Changes the locale lists are shown in. */
  setLocale(locale: FormatLocale | undefined): void {
    this.options.locale = locale;
  }

  /** The rule that applies to a cell, with its ranges and formulas where they are now. */
  ruleAt(row: number, col: number): DataValidation | undefined {
    const entry = this.entryAt(row, col);
    if (!entry) return undefined;
    const rule: DataValidation = { ...entry.rule, ranges: entry.ranges };
    if (entry.formula1 !== undefined) rule.formula1 = entry.formula1;
    if (entry.formula2 !== undefined) rule.formula2 = entry.formula2;
    return rule;
  }

  /** The message a cell shows while it is selected; undefined when it has none. */
  prompt(row: number, col: number): { title: string; text: string } | undefined {
    const entry = this.entryAt(row, col);
    if (!entry || !entry.rule.showInputMessage) return undefined;
    const title = typeof entry.rule.promptTitle === 'string' ? entry.rule.promptTitle : '';
    const text = typeof entry.rule.prompt === 'string' ? entry.rule.prompt : '';
    return title === '' && text === '' ? undefined : { title, text };
  }

  /** Whether a cell shows the button of a list: a list rule that does not hide it, with a source that can be read. */
  hasList(row: number, col: number): boolean {
    const entry = this.entryAt(row, col);
    if (!entry || entry.rule.type !== 'list' || entry.rule.hideDropDown) return false;
    try {
      const source = this.source(entry, row, col);
      if (!source || source.kind !== 'formula') return source !== undefined;
      // A formula has a list to offer when the engine can count its values.
      const engine = this.engine();
      if (!engine) return false;
      const count = engine.evaluateShifted(`ROWS(${source.text})*COLUMNS(${source.text})`, { sheet: this.sheetIndex(), row: entry.anchorRow, col: entry.anchorCol }, { row, col });
      return typeof count === 'number' && count >= 1;
    } catch (error) {
      return false;
    }
  }

  /** The values a cell offers; undefined when its rule is not a list or the source cannot be read. */
  list(row: number, col: number): ValidationList | undefined {
    try {
      const entry = this.entryAt(row, col);
      if (!entry || entry.rule.type !== 'list') return undefined;
      const items: string[] = [];
      let truncated = false;
      const known = this.eachSource(entry, row, col, (_value, text) => {
        if (items.length >= MAX_LIST_ITEMS) {
          truncated = true;
          return true;
        }
        const shown = text();
        if (shown !== '') items.push(shown);
        return false;
      });
      return known === 'unknown' ? undefined : { items, truncated: truncated || known === 'cut' };
    } catch (error) {
      return undefined;
    }
  }

  /**
   * Whether a cell may hold what is about to be written into it. `cell` is the cell the
   * typed text turns into (undefined: the cell is emptied) and `typed` the text itself.
   * A custom rule reads the cell through its formula, and maybe formulas that read the cell,
   * so it needs the new cell in place and those calculated: `tryOut` runs a function while
   * they are, and puts everything back (see tryOps()). When it answers undefined the value
   * could not be tried, and the rule is not applied.
   *
   * A formula is judged by what it gives now, where it is entered; it is not judged again
   * when the cells it reads change. A formula that cannot be calculated here (no engine, a
   * function the engine lacks, another workbook) is let through, since Excel may well know
   * what it gives. UNSURE: so is a formula that gives an error value; whether Excel refuses
   * those is not known, and refusing what Excel takes would lose the input.
   *
   * A cell whose rule says nothing happens on a wrong value is not checked. Never throws.
   */
  check(row: number, col: number, cell: Cell | undefined, typed: string, tryOut?: <T>(run: () => T) => T | undefined): ValidationVerdict {
    try {
      const entry = this.entryAt(row, col);
      if (!entry) return OK;
      const rule = entry.rule;
      if (!rule.showErrorMessage || rule.type === 'none') return OK;
      const date1904 = this.workbook.date1904 === true;
      let value = plainValue(cell, date1904);
      if (cell && cell.formula) {
        const result = this.resultOf(cell.formula.text, row, col);
        if (result === undefined) return OK;
        value = result;
        // A list looks for the text that was typed; of a formula that is what it gives.
        typed = typeof result === 'string' ? result : typeof result === 'boolean' ? (result ? 'TRUE' : 'FALSE') : String(result);
      }
      const failed: ValidationVerdict = { ok: false, style: rule.errorStyle };
      if (typeof rule.errorTitle === 'string' && rule.errorTitle !== '') failed.title = rule.errorTitle;
      if (typeof rule.error === 'string' && rule.error !== '') failed.message = rule.error;
      if (value === undefined || value === '') {
        // An error value typed in is never what a rule asks for; an empty cell is, where blanks are allowed.
        if (cell && cell.kind === 'error' && !cell.formula) return failed;
        return rule.allowBlank ? OK : failed;
      }
      let passes: boolean | undefined;
      switch (rule.type) {
        case 'whole':
          passes = typeof value === 'number' && Math.floor(value) === value ? this.inLimits(entry, value, row, col) : false;
          break;
        case 'decimal':
        case 'date':
        case 'time':
          passes = typeof value === 'number' ? this.inLimits(entry, value, row, col) : false;
          break;
        case 'textLength': {
          const text = typeof value === 'string' ? value : typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : formatValue(value, 'General').text;
          passes = this.inLimits(entry, text.length, row, col);
          break;
        }
        case 'list':
          passes = this.inList(entry, value, typed, row, col);
          break;
        case 'custom': {
          const formula = entry.formula1;
          if (formula === undefined) break;
          const evaluate = () => this.truth(formula, entry, row, col);
          passes = tryOut ? tryOut(evaluate) : evaluate();
          break;
        }
        default:
        // A kind of rule this engine does not know.
      }
      // undefined: the rule could not be worked out, and nothing is refused for that.
      return passes === false ? failed : OK;
    } catch (error) {
      return OK;
    }
  }

  // -------------------------------------------------------------------------

  private engine(): CalcEngine | undefined {
    try {
      return this.options.engine ? this.options.engine() : undefined;
    } catch (error) {
      return undefined;
    }
  }

  private entryAt(row: number, col: number): Entry | undefined {
    const entries = this.stale() ? this.build() : (this.entries as Entry[]);
    for (let i = 0; i < entries.length; i++) if (inRanges(entries[i].ranges, row, col)) return entries[i];
    return undefined;
  }

  /** Rows or columns were inserted or deleted, or the rules replaced, since the entries were built. */
  private stale(): boolean {
    if (!this.entries || this.rulesSeen !== this.sheet.dataValidations) return true;
    const sheets = this.workbook.sheets || [];
    if (sheets.length !== this.editsSeen.length || this.removedSeen !== this.workbook.removedSheetNames) return true;
    for (let i = 0; i < sheets.length; i++) {
      // A sheet inserted or deleted moves the others in the list; a renamed one keeps its place.
      if (sheets[i] !== this.sheetsSeen[i] || sheets[i].name !== this.namesSeen[i] || sheets[i].structuralEdits !== this.editsSeen[i]) return true;
    }
    return false;
  }

  private build(): Entry[] {
    const sheet = this.sheet;
    const sheets = this.workbook.sheets || [];
    this.rulesSeen = sheet.dataValidations;
    this.editsSeen = sheets.map((other) => other.structuralEdits);
    this.sheetsSeen = sheets.slice();
    this.namesSeen = sheets.map((other) => other.name);
    this.removedSeen = this.workbook.removedSheetNames;
    // The formulas of the rules are as the file has them and name sheets as the file does: the
    // edits are matched with those names, and only then are the sheets given the names they have now.
    const fileName = fileSheetName(sheet);
    const ownName = fileName === undefined ? NOT_IN_FILE : fileName;
    const nameMap = sheetNameMap(this.workbook);
    const own: StructuralEdit[] = (sheet.structuralEdits || []).map((edit) => ({ ...edit, sheetName: ownName }));
    const foreign: StructuralEdit[] = [];
    for (const other of sheets) {
      if (other === sheet || !other.structuralEdits) continue;
      // Nothing in the file refers to a sheet that is not in it.
      const name = fileSheetName(other);
      if (name === undefined) continue;
      for (const edit of other.structuralEdits) foreign.push({ ...edit, sheetName: name });
    }
    const entries: Entry[] = [];
    const rules = Array.isArray(sheet.dataValidations) ? sheet.dataValidations : [];
    for (const rule of rules) {
      if (!rule || !Array.isArray(rule.ranges)) continue;
      let ranges = rule.ranges.filter((range) => !!range && range.top <= range.bottom && range.left <= range.right);
      if (ranges.length === 0) continue;
      const texts = [rule.formula1, rule.formula2].map((text) => (typeof text === 'string' ? text : undefined));
      // The formulas are written for the top-left cell of the first range; they follow that corner.
      const adjust = (edit: StructuralEdit, from: CellRange, to: CellRange) => {
        for (let i = 0; i < texts.length; i++) {
          const text = texts[i];
          if (text !== undefined) texts[i] = adjustAnchoredFormula(text, ownName, edit, { row: from.top, col: from.left }, { row: to.top, col: to.left }, true);
        }
      };
      for (const edit of own) {
        const moved: CellRange[] = [];
        for (const range of ranges) {
          const next = adjustRange(range, edit);
          if (next) moved.push(next);
        }
        if (moved.length === 0) {
          ranges = moved;
          break;
        }
        adjust(edit, ranges[0], moved[0]);
        ranges = moved;
      }
      if (ranges.length === 0) continue;
      for (const edit of foreign) adjust(edit, ranges[0], ranges[0]);
      if (nameMap) for (let i = 0; i < texts.length; i++) if (texts[i] !== undefined) texts[i] = renameSheetsInFormula(texts[i] as string, nameMap);
      entries.push({ rule, ranges, anchorRow: ranges[0].top, anchorCol: ranges[0].left, formula1: texts[0], formula2: texts[1] });
    }
    this.entries = entries;
    return entries;
  }

  /**
   * What the formula of a rule gives at a cell; undefined when it cannot be calculated: there
   * is no engine, the text is not a formula, it uses what the engine does not know (a function
   * it lacks, LET, a name that is not defined: #NAME?) or cannot reach (another workbook).
   * Any other error is a result like any other.
   */
  private calculate(formula: string, entry: Entry, row: number, col: number): Scalar | undefined {
    const engine = this.engine();
    if (!engine) return undefined;
    let node: Node;
    try {
      node = parseFormula(formula);
    } catch (error) {
      return undefined;
    }
    if (isBeyondEngine(node)) return undefined;
    const value = engine.evaluateShifted(formula, { sheet: this.sheetIndex(), row: entry.anchorRow, col: entry.anchorCol }, { row, col });
    return isError(value) && value.error === '#NAME?' ? undefined : value;
  }

  /**
   * What a formula typed into a cell gives there, from the values the cells hold now;
   * undefined when that is not known (see calculate()) or is an error value. Not known
   * either: while a calculation is under way (`incomplete`), when cells the formula reads
   * may still hold values of before - the caller finishes the calculation first where it
   * can; and of a formula that names its own cell, alone or within a range, which would be
   * given the value the cell had (to Excel that is a circular reference).
   */
  private resultOf(formula: string, row: number, col: number): Plain | undefined {
    const engine = this.engine();
    if (!engine || engine.incomplete || typeof formula !== 'string' || formula === '') return undefined;
    let node: Node;
    try {
      node = parseFormula(formula);
    } catch (error) {
      return undefined;
    }
    if (isBeyondEngine(node) || this.readsCell(node, row, col)) return undefined;
    const value = engine.evaluateText(formula, { sheet: this.sheetIndex(), row, col });
    return value === null || isError(value) ? undefined : value;
  }

  /** Whether a formula names a cell of this sheet in a reference of its own (not through a name or a function that makes references). */
  private readsCell(root: Node, row: number, col: number): boolean {
    const own = this.sheet.name.toLowerCase();
    const stack: Node[] = [root];
    while (stack.length > 0) {
      const node = stack.pop() as Node;
      switch (node.type) {
        case 'ref': {
          if (node.external || node.sheetEnd !== undefined || (node.sheet !== undefined && node.sheet.toLowerCase() !== own)) break;
          const range = node.range;
          if (row >= range.top && row <= range.bottom && col >= range.left && col <= range.right) return true;
          break;
        }
        case 'call':
          for (const arg of node.args) if (arg) stack.push(arg);
          break;
        case 'unary':
          stack.push(node.operand);
          break;
        case 'binary':
          stack.push(node.left, node.right);
          break;
        default:
        // Constants and names hold nothing more.
      }
    }
    return false;
  }

  /**
   * The number a limit of a rule stands for at a cell; undefined when it cannot be worked
   * out (the rule is then not applied), NO_LIMIT when its formula gives an error.
   */
  private limit(text: string | undefined, entry: Entry, row: number, col: number): number | typeof NO_LIMIT | undefined {
    if (text === undefined || text.trim() === '') return undefined;
    const constant = plainNumber(text);
    if (constant !== undefined) return constant;
    const value = this.calculate(text, entry, row, col);
    if (isError(value)) return NO_LIMIT;
    return typeof value === 'number' ? value : typeof value === 'boolean' ? (value ? 1 : 0) : undefined;
  }

  private inLimits(entry: Entry, value: number, row: number, col: number): boolean | undefined {
    const a = this.limit(entry.formula1, entry, row, col);
    const operator = entry.rule.operator || 'between';
    const two = operator === 'between' || operator === 'notBetween';
    const b = two ? this.limit(entry.formula2, entry, row, col) : undefined;
    // A limit that cannot be worked out lifts the rule, whatever the other limit is.
    if (a === undefined || (two && b === undefined)) return undefined;
    // UNSURE: Excel says of such a rule that it "currently evaluates to an error"; like a custom
    // formula that gives an error, it is taken to refuse what is typed.
    if (a === NO_LIMIT || b === NO_LIMIT) return false;
    return compare(value, operator, a, b);
  }

  /** What a custom formula says of the cell: any number but 0, or TRUE, lets the value pass. */
  private truth(formula: string, entry: Entry, row: number, col: number): boolean | undefined {
    const value = this.calculate(formula, entry, row, col);
    if (value === undefined) return undefined;
    if (typeof value === 'boolean') return value;
    if (typeof value === 'number') return value !== 0;
    // Text or an error: the formula does not say yes.
    return false;
  }

  private sheetIndex(): number {
    const index = (this.workbook.sheets || []).indexOf(this.sheet);
    return index < 0 ? 0 : index;
  }

  /** A tree node that names cells of one sheet, as the range it names for the cell (row, col). */
  private cellsOf(node: Node, home: Sheet, entry: Entry | undefined, row: number, col: number): { sheet: Sheet; range: CellRange } | undefined {
    if (node.type !== 'ref' || node.external || node.sheetEnd !== undefined) return undefined;
    let sheet = home;
    if (node.sheet !== undefined) {
      const wanted = node.sheet.toLowerCase();
      const found = (this.workbook.sheets || []).find((other) => other.name.toLowerCase() === wanted);
      if (!found) return undefined;
      sheet = found;
    }
    // Relative sides are as far from the cell as they are, in the rule, from its first cell.
    const dRow = entry ? row - entry.anchorRow : 0;
    const dCol = entry ? col - entry.anchorCol : 0;
    const abs = node.abs;
    const top = node.wholeCols ? 0 : node.range.top + (abs.top ? 0 : dRow);
    const bottom = node.wholeCols ? MAX_ROWS - 1 : node.range.bottom + (abs.bottom ? 0 : dRow);
    const left = node.wholeRows ? 0 : node.range.left + (abs.left ? 0 : dCol);
    const right = node.wholeRows ? MAX_COLS - 1 : node.range.right + (abs.right ? 0 : dCol);
    if (!(top >= 0 && left >= 0 && bottom < MAX_ROWS && right < MAX_COLS && top <= bottom && left <= right)) return undefined;
    return { sheet, range: { top, left, bottom, right } };
  }

  /** Where the list of a rule takes its values from, for one cell; undefined when the rule states nothing usable. */
  private source(entry: Entry, row: number, col: number): Source | undefined {
    const formula = entry.formula1;
    if (formula === undefined || formula.trim() === '') return undefined;
    const constant = constantItems(formula);
    if (constant) return { kind: 'constant', items: constant };
    let node: Node;
    try {
      node = parseFormula(formula);
    } catch (error) {
      return undefined;
    }
    const direct = this.cellsOf(node, this.sheet, entry, row, col);
    if (direct) return { kind: 'cells', ...direct };
    // A defined name: the sheet's own before the workbook's; names that name names are followed a few steps.
    let home = this.sheet;
    for (let depth = 0; node.type === 'name' && depth < MAX_NAME_DEPTH; depth++) {
      const wanted = node.name.toLowerCase();
      const scope = node.sheet === undefined ? this.sheetIndex() : (this.workbook.sheets || []).findIndex((other) => other.name.toLowerCase() === (node as { sheet: string }).sheet.toLowerCase());
      const names = this.workbook.definedNames || [];
      const found = names.find((name) => name.localSheet === scope && name.name.toLowerCase() === wanted) || names.find((name) => name.localSheet === undefined && name.name.toLowerCase() === wanted);
      if (!found) return undefined;
      try {
        node = parseFormula(found.value.charAt(0) === '=' ? found.value.slice(1) : found.value);
      } catch (error) {
        return undefined;
      }
      if (found.localSheet !== undefined && this.workbook.sheets[found.localSheet]) home = this.workbook.sheets[found.localSheet];
      // A name stands for the same cells wherever it is used.
      const cells = this.cellsOf(node, home, undefined, row, col);
      if (cells) return { kind: 'cells', ...cells, named: true };
    }
    return node.type === 'name' || node.type === 'ref' ? undefined : { kind: 'formula', text: formula };
  }

  /**
   * Calls `visit` with every value of the list of a cell, in order, until it returns true.
   * `text` gives the value as its cell shows it. Returns "unknown" when the source cannot
   * be read, "cut" when it was too large to be read to its end.
   */
  private eachSource(entry: Entry, row: number, col: number, visit: (value: Plain, text: () => string) => boolean): 'done' | 'cut' | 'unknown' {
    const source = this.source(entry, row, col);
    if (!source) return 'unknown';
    const locale = this.options.locale;
    const date1904 = this.workbook.date1904 === true;
    if (source.kind === 'constant') {
      for (const item of source.items) if (visit(item, () => item)) break;
      return 'done';
    }
    if (source.kind === 'cells') {
      const { sheet, range } = source;
      const bottom = Math.min(range.bottom, sheet.rowCount - 1);
      const right = Math.min(range.right, sheet.colCount - 1);
      if ((bottom - range.top + 1) * (right - range.left + 1) > MAX_SOURCE_CELLS) return 'unknown';
      const styles = this.workbook.styles;
      for (let r = range.top; r <= bottom; r++) {
        const data = sheet.rows.get(r);
        if (!data) continue;
        for (let c = range.left; c <= right && c < data.cells.length; c++) {
          const cell = data.cells[c];
          const value = plainValue(cell, date1904);
          if (value === undefined || value === '') continue;
          const text = () => {
            const style = resolveCellStyle(styles, cell.style);
            return formatCell(cell, style.numFmt, { locale, date1904, builtinId: builtinFormatId(styles, style.numFmtId) }).text;
          };
          if (visit(value, text)) return 'done';
        }
      }
      return 'done';
    }
    // Any other formula (OFFSET, INDIRECT, ...): the engine gives all of its values at once.
    const engine = this.engine();
    if (!engine) return 'unknown';
    const found = engine.evaluateRange(source.text, { sheet: this.sheetIndex(), row, col }, { anchor: { row: entry.anchorRow, col: entry.anchorCol }, maxCells: MAX_SOURCE_CELLS });
    if (!('values' in found) || found.values.length === 0) return 'unknown';
    for (const value of found.values) {
      if (value === null || value === '' || typeof value === 'object') continue;
      const plain = value as Plain;
      if (visit(plain, () => (typeof plain === 'string' ? plain : formatValue(plain, 'General', { locale }).text))) return 'done';
    }
    return found.truncated ? 'cut' : 'done';
  }

  private inList(entry: Entry, value: Plain, typed: string, row: number, col: number): boolean | undefined {
    const source = this.source(entry, row, col);
    if (!source) return undefined;
    const text = typeof value === 'string' ? value : typed.trim();
    let found = false;
    if (source.kind === 'constant') {
      // UNSURE: Excel compares a typed value with the items of a list written into the rule by case,
      // and with the cells of a range without regard to it.
      const known = this.eachSource(entry, row, col, (item) => {
        const number = typeof value === 'number' ? plainNumber(String(item)) : undefined;
        found = item === text || (number !== undefined && number === value) || (typeof value === 'boolean' && String(item).toUpperCase() === (value ? 'TRUE' : 'FALSE'));
        return found;
      });
      return known === 'unknown' ? undefined : found;
    }
    const locale = this.options.locale;
    const folded = typeof value === 'string' ? foldCase(value, locale) : undefined;
    let filled = 0;
    const known = this.eachSource(entry, row, col, (item) => {
      filled++;
      found = typeof item === 'string' ? folded !== undefined && foldCase(item, locale) === folded : item === value;
      return found;
    });
    if (found) return true;
    // A source that could not be read to its end may hold the value further down.
    if (known !== 'done') return undefined;
    // UNSURE: a list given by a defined name that has empty cells in it lets every value through
    // where the rule ignores blanks; Excel is described to do that for names, not for cells
    // written into the rule.
    if (source.kind === 'cells' && source.named && entry.rule.allowBlank) {
      const { top, left, bottom, right } = source.range;
      if (filled < (bottom - top + 1) * (right - left + 1)) return true;
    }
    return false;
  }
}

/** Most calls a trial gives the engine to finish a calculation that runs out of its budget. */
const MAX_TRIAL_SLICES = 16;

/**
 * Remembers the rows the sheets had when it was made, so that they can be put back: the
 * very same objects, whatever replaced them in between. Rows are never changed in place,
 * only replaced through the map of their sheet (by operations and by the engine alike), so
 * it is enough to watch the maps; watching costs by the rows that change, not by the rows
 * there are.
 */
class RowJournal {
  private readonly watched: { rows: Map<number, Row>; was: Map<number, Row | undefined> }[] = [];

  constructor(sheets: Sheet[]) {
    for (const sheet of sheets) {
      const rows = sheet.rows;
      const was = new Map<number, Row | undefined>();
      const note = (key: number) => {
        if (!was.has(key)) was.set(key, Map.prototype.get.call(rows, key) as Row | undefined);
      };
      rows.set = function (key: number, value: Row) {
        note(key);
        return Map.prototype.set.call(this, key, value) as Map<number, Row>;
      };
      rows.delete = function (key: number) {
        note(key);
        return Map.prototype.delete.call(this, key);
      };
      rows.clear = function () {
        Map.prototype.forEach.call(this, (_row: Row, key: number) => note(key));
        Map.prototype.clear.call(this);
      };
      this.watched.push({ rows, was });
    }
  }

  /** Puts back every row that was replaced, added or removed since the journal was made or last rolled back. */
  rollback(): void {
    for (const { rows, was } of this.watched) {
      was.forEach((row, key) => {
        if (row === undefined) Map.prototype.delete.call(rows, key);
        else Map.prototype.set.call(rows, key, row);
      });
      was.clear();
    }
  }

  /** Stops watching; the maps are plain maps again. */
  close(): void {
    for (const { rows } of this.watched) {
      const own = rows as unknown as { set?: unknown; delete?: unknown; clear?: unknown };
      delete own.set;
      delete own.delete;
      delete own.clear;
    }
    this.watched.length = 0;
  }
}

const addressKey = (cell: CellAddress) => `${cell.sheet}:${cell.row}:${cell.col}`;

/** Recalculates until the engine is done; false when it is not done after MAX_TRIAL_SLICES calls. */
function settle(engine: CalcEngine, changed: CellAddress[]): boolean {
  engine.recalculate(changed);
  for (let i = 1; engine.incomplete && i < MAX_TRIAL_SLICES; i++) engine.recalculate([]);
  return !engine.incomplete;
}

/**
 * Runs `run` as if operations had been applied: the cells are written and every formula
 * that depends on them, on any sheet, is calculated, so that a rule that reads such a formula
 * judges the value that is being typed and not the one before it. Afterwards everything is
 * taken back: the sheets hold the very rows they held (cells, values of the file, `modified`,
 * `changedRows`, used extent), and the engine lists as differing from the file what it
 * listed before. Nothing is drawn or reported in between: this runs within one call.
 *
 * Only changes of cells can be tried. Answers undefined, without calling `run`, when the
 * engine could not finish calculating within its budget several times over: there is then
 * nothing to judge the value by. The same when the engine is still in the middle of a
 * calculation (`incomplete`): what that calculation writes is real and must not be taken
 * back with the trial, so the caller finishes it first. Without an engine the cells are put
 * in place and taken back, and nothing is calculated.
 *
 * UNSURE: a formula that was over the engine's budget before and is calculated during the
 * trial is no longer listed as stale afterwards, although its cell has its old value again.
 */
export function tryOps<T>(workbook: Workbook, engine: CalcEngine | undefined, ops: Op[], run: () => T): T | undefined {
  const sheets = workbook.sheets || [];
  for (const op of ops) if (op.type !== 'setCells' || !sheets[op.sheet]) return run();
  if (engine && engine.incomplete) return undefined;
  const saved = sheets.map((sheet) => ({ modified: sheet.modified, changedRows: sheet.changedRows, rowCount: sheet.rowCount, colCount: sheet.colCount }));
  // What the trial marks as changed must not end up in the sets the writer reads.
  for (const sheet of sheets) if (sheet.changedRows) sheet.changedRows = new Set(sheet.changedRows);
  const mismatched = engine ? engine.mismatched : [];
  const circular = engine ? engine.circular.slice() : [];
  const journal = new RowJournal(sheets);
  const changed: CellAddress[] = [];
  // A formula that is replaced or put in place changes the graph of the engine, which then has to be told again.
  let formulas = false;
  try {
    const inverse: Op[] = [];
    try {
      for (const op of ops) {
        if (op.type !== 'setCells') continue;
        const undo = applyOp(workbook, op);
        inverse.push(undo);
        if (undo.type !== 'setCells') continue;
        // The inverse also names the cells the operation changed on its own (dependants of a shared formula).
        for (const change of undo.cells) {
          changed.push({ sheet: op.sheet, row: change.row, col: change.col });
          const now = getCell(sheets[op.sheet], change.row, change.col);
          if ((change.cell && change.cell.formula) || (now && now.formula)) formulas = true;
        }
      }
      if (engine && !settle(engine, changed)) return undefined;
      return run();
    } finally {
      for (let i = inverse.length - 1; i >= 0; i--) applyOp(workbook, inverse[i]);
      journal.rollback();
    }
  } finally {
    try {
      if (engine) {
        // The engine reads values from the cells, which are back; what it keeps beside them is brought back here.
        if (formulas || engine.incomplete) {
          settle(engine, changed);
          journal.rollback();
        }
        // A formula whose value of the file stood although the engine calculates another one was
        // calculated during the trial as one that changed. Named again, with nothing changed, it
        // is found to differ again; a circle met only while calculating is met again the same way.
        const still = new Set(engine.mismatched.map(addressKey));
        const lost = mismatched.filter((cell) => !still.has(addressKey(cell)));
        const circles = new Set(engine.circular.map(addressKey));
        for (const cell of circular) if (!circles.has(addressKey(cell))) lost.push(cell);
        if (lost.length > 0) {
          settle(engine, lost);
          journal.rollback();
        }
      }
    } catch (error) {
      // The engine never throws; if it did, the cells are still put back below.
      journal.rollback();
    }
    journal.close();
    sheets.forEach((sheet, index) => {
      const was = saved[index];
      if (was.modified === undefined) delete sheet.modified;
      else sheet.modified = was.modified;
      if (was.changedRows) sheet.changedRows = was.changedRows;
      else delete sheet.changedRows;
      sheet.rowCount = was.rowCount;
      sheet.colCount = was.colCount;
    });
  }
}
