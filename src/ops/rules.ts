// Data validations and conditional formats as the editor changes them: a rule
// set on a range, changed or taken away. Like the other commands these functions
// only build operations; the caller applies them through a History so that each
// is one undo step.
//
// Sheet.dataValidations and Sheet.conditionalFormats hold the rules of the file
// the way the file has them: in the rows and columns of the file, naming sheets as
// the file does. The rows and columns moved since (Sheet.structuralEdits) and the
// sheets renamed or deleted since (sheetNameMap()) are applied where such a rule
// is looked at. A rule that is added or changed here is `live`: its ranges and
// formulas are those of the workbook as it is now, and the operations that move
// cells or name sheets keep them so (see liveRulesMoved() and liveRulesRenamed(),
// which applyOp() calls). dataValidationNow() and conditionalFormatNow() are the
// one rule for where either kind applies and what its formulas say: the
// evaluation, the writer and the builder of a new file all ask them.

import { parseFormula } from '../calc/parse';
import { adjustAnchoredFormula, adjustFormula, adjustRange, formulaNamesSheetsUnread, renameSheetsInFormula, shiftFormula, type StructuralEdit } from '../calc/refs';
import { textToNumber } from '../calc/values';
import { resolveColor } from '../model/sheet/color';
import {
  MAX_COLS,
  MAX_ROWS,
  fileSheetName,
  rangeName,
  sheetNameMap,
  type CellRange,
  type Color,
  type ConditionalFormat,
  type ConditionalRule,
  type ConditionalValue,
  type DataValidation,
  type DifferentialFormat,
  type Sheet,
  type SheetNameMap,
  type StyleSheet,
  type Workbook,
} from '../model/sheet/types';
import type { Op } from './ops';

// ---------------------------------------------------------------------------
// Ranges

const validRange = (range: CellRange | undefined): range is CellRange =>
  !!range &&
  Math.floor(range.top) === range.top &&
  Math.floor(range.left) === range.left &&
  Math.floor(range.bottom) === range.bottom &&
  Math.floor(range.right) === range.right &&
  range.top >= 0 &&
  range.left >= 0 &&
  range.top <= range.bottom &&
  range.left <= range.right &&
  range.bottom < MAX_ROWS &&
  range.right < MAX_COLS;

const intersects = (a: CellRange, b: CellRange) => a.top <= b.bottom && b.top <= a.bottom && a.left <= b.right && b.left <= a.right;

export const rangesIntersect = (a: readonly CellRange[], b: readonly CellRange[]): boolean => a.some((x) => b.some((y) => intersects(x, y)));

/**
 * What is left of a range once the cells of `cut` are taken out of it: at most four
 * rectangles that do not overlap (the rows above the cut, the rows below it, and what is
 * left and right of it in between), in that order.
 */
export function subtractRange(range: CellRange, cut: CellRange): CellRange[] {
  if (!intersects(range, cut)) return [range];
  const out: CellRange[] = [];
  const top = Math.max(range.top, cut.top);
  const bottom = Math.min(range.bottom, cut.bottom);
  if (cut.top > range.top) out.push({ top: range.top, left: range.left, bottom: cut.top - 1, right: range.right });
  if (cut.bottom < range.bottom) out.push({ top: cut.bottom + 1, left: range.left, bottom: range.bottom, right: range.right });
  if (cut.left > range.left) out.push({ top, left: range.left, bottom, right: cut.left - 1 });
  if (cut.right < range.right) out.push({ top, left: cut.right + 1, bottom, right: range.right });
  return out;
}

/** Ranges of one rule after cutting: more than any rule of a file has, and a bound on what one command makes. */
export const MAX_RULE_RANGES = 8192;

/** `ranges` without the cells of `cuts`; the order of what is left follows the order of `ranges`. */
export function subtractRanges(ranges: readonly CellRange[], cuts: readonly CellRange[]): CellRange[] {
  let left = ranges.slice();
  for (const cut of cuts) {
    const next: CellRange[] = [];
    for (const range of left) {
      for (const piece of subtractRange(range, cut)) next.push(piece);
      if (next.length > MAX_RULE_RANGES) throw new RangeError('The rule would be left with too many separate ranges');
    }
    left = next;
    if (left.length === 0) break;
  }
  return left;
}

/** The ranges a command is given: inside the sheet, none of them overlapping another. */
function targetRanges(ranges: CellRange | readonly CellRange[]): CellRange[] {
  const list = Array.isArray(ranges) ? (ranges as readonly CellRange[]) : [ranges as CellRange];
  if (list.length === 0 || list.length > MAX_RULE_RANGES) throw new RangeError('A rule applies to at least one range');
  let out: CellRange[] = [];
  for (const range of list) {
    if (!range || typeof range !== 'object') throw new RangeError('A range is outside the sheet');
    const tidy = { top: Math.min(range.top, range.bottom), left: Math.min(range.left, range.right), bottom: Math.max(range.top, range.bottom), right: Math.max(range.left, range.right) };
    if (!validRange(tidy)) throw new RangeError('A range is outside the sheet');
    // Cells named twice are in the rule once.
    out = out.concat(subtractRanges([tidy], out));
  }
  return out;
}

const sameRanges = (a: readonly CellRange[], b: readonly CellRange[]) =>
  a.length === b.length && a.every((range, i) => range.top === b[i].top && range.left === b[i].left && range.bottom === b[i].bottom && range.right === b[i].right);

// ---------------------------------------------------------------------------
// Where a rule of the file is now

/**
 * Stands for the name in the file of a sheet that is not in the file: no reference of the
 * file names it (a sheet name holds no control character), references without a sheet name
 * still follow the sheet's own edits.
 */
const NOT_IN_FILE = '\u0000';

/** What a rule kept the way the file has it is read through: the edits since, under the names of the file, and what those names stand for now. */
export interface RuleSpace {
  /** The name the file has for the sheet of the rules. */
  ownName: string;
  /** Rows and columns moved in the sheet itself, in order. */
  own: StructuralEdit[];
  /** The same for the other sheets of the file. */
  foreign: StructuralEdit[];
  nameMap: SheetNameMap | undefined;
}

/** The space the rules of the file of a sheet are read through, for the workbook as it is at this moment. */
export function ruleSpace(workbook: Workbook, sheet: Sheet): RuleSpace {
  // The formulas of the rules are as the file has them and name sheets as the file does: the
  // edits are matched with those names, and only then are the sheets given the names they have now.
  const fileName = fileSheetName(sheet);
  const ownName = fileName === undefined ? NOT_IN_FILE : fileName;
  const own: StructuralEdit[] = (sheet.structuralEdits || []).map((edit) => ({ ...edit, sheetName: ownName }));
  const foreign: StructuralEdit[] = [];
  for (const other of workbook.sheets || []) {
    if (other === sheet || !other.structuralEdits) continue;
    // Nothing in the file refers to a sheet that is not in it.
    const name = fileSheetName(other);
    if (name === undefined) continue;
    for (const edit of other.structuralEdits) foreign.push({ ...edit, sheetName: name });
  }
  return { ownName, own, foreign, nameMap: sheetNameMap(workbook) };
}

/**
 * Ranges and formulas of the file as they are now. The formulas are written for the top-left
 * cell of the first range and follow that corner; relative references go round the edge of
 * the sheet, as they do when a rule is evaluated. `plain` formulas (thresholds) stand on
 * their own. Undefined when all the cells of the ranges were deleted.
 */
export function movedThroughSpace(
  ranges: readonly CellRange[],
  texts: string[],
  space: RuleSpace,
): CellRange[] | undefined {
  let now = ranges.filter((range) => !!range && range.top <= range.bottom && range.left <= range.right);
  if (now.length === 0) return undefined;
  const adjust = (edit: StructuralEdit, from: CellRange, to: CellRange) => {
    const before = { row: from.top, col: from.left };
    const after = { row: to.top, col: to.left };
    for (let i = 0; i < texts.length; i++) texts[i] = adjustAnchoredFormula(texts[i], space.ownName, edit, before, after, true);
  };
  for (const edit of space.own) {
    const moved: CellRange[] = [];
    for (const range of now) {
      const next = adjustRange(range, edit);
      if (next) moved.push(next);
    }
    if (moved.length === 0) return undefined;
    adjust(edit, now[0], moved[0]);
    now = moved;
  }
  for (const edit of space.foreign) adjust(edit, now[0], now[0]);
  if (space.nameMap) for (let i = 0; i < texts.length; i++) texts[i] = renameSheetsInFormula(texts[i], space.nameMap);
  return now;
}

/**
 * A data validation as it applies now: its ranges in the rows and columns of the sheet as it
 * is, its formulas for the top-left cell of the first of them, under the names the sheets
 * have now. A live rule is that already and is handed back as it is; a rule of the file is
 * read through `space`. Undefined for a rule none of whose cells is left.
 */
export function dataValidationNow(rule: DataValidation, space: RuleSpace): DataValidation | undefined {
  if (!rule || !Array.isArray(rule.ranges)) return undefined;
  if (rule.live) return rule.ranges.length > 0 ? rule : undefined;
  const slots: ('formula1' | 'formula2')[] = [];
  const texts: string[] = [];
  for (const name of ['formula1', 'formula2'] as const) {
    if (typeof rule[name] !== 'string') continue;
    slots.push(name);
    texts.push(rule[name] as string);
  }
  const ranges = movedThroughSpace(rule.ranges, texts, space);
  if (!ranges) return undefined;
  const now: DataValidation = { ...rule, ranges };
  slots.forEach((name, i) => {
    now[name] = texts[i];
  });
  return now;
}

/** A threshold that is a formula (not a number) after the edits of the space; the value itself when it is none or did not change. */
function thresholdsNow(values: ConditionalValue[], space: RuleSpace): ConditionalValue[] {
  return values.map((value) => {
    if (!value || typeof value.value !== 'string' || textToNumber(value.value) !== undefined) return value;
    let text = value.value;
    for (const edit of space.own) text = adjustFormula(text, space.ownName, edit);
    for (const edit of space.foreign) text = adjustFormula(text, space.ownName, edit);
    if (space.nameMap) text = renameSheetsInFormula(text, space.nameMap);
    return text === value.value ? value : { ...value, value: text };
  });
}

/** A rule whose thresholds are formulas, with those formulas as `change` makes them; the rule itself when none changed. */
function withThresholds(rule: ConditionalRule, change: (values: ConditionalValue[]) => ConditionalValue[]): ConditionalRule {
  const same = (a: ConditionalValue[], b: ConditionalValue[]) => a.every((value, i) => value === b[i]);
  if (rule.colorScale && Array.isArray(rule.colorScale.values)) {
    const values = change(rule.colorScale.values);
    return same(values, rule.colorScale.values) ? rule : { ...rule, colorScale: { ...rule.colorScale, values } };
  }
  if (rule.iconSet && Array.isArray(rule.iconSet.values)) {
    const values = change(rule.iconSet.values);
    return same(values, rule.iconSet.values) ? rule : { ...rule, iconSet: { ...rule.iconSet, values } };
  }
  if (rule.dataBar) {
    const [min, max] = change([rule.dataBar.min, rule.dataBar.max]);
    return min === rule.dataBar.min && max === rule.dataBar.max ? rule : { ...rule, dataBar: { ...rule.dataBar, min, max } };
  }
  return rule;
}

/** A rule with every formula of it (its own, and the thresholds that are formulas) as `change` makes it; the rule itself when none changed. */
export function ruleWithFormulas(rule: ConditionalRule, change: (text: string) => string): ConditionalRule {
  let next = rule;
  const formulas = rule.formulas.map(change);
  if (formulas.some((text, i) => text !== rule.formulas[i])) next = { ...next, formulas };
  return withThresholds(next, (values) =>
    values.map((value) => {
      if (!value || typeof value.value !== 'string' || textToNumber(value.value) !== undefined) return value;
      const text = change(value.value);
      return text === value.value ? value : { ...value, value: text };
    }),
  );
}

/**
 * A conditional format as it applies now: its ranges in the rows and columns of the sheet as
 * it is, the formulas of its rules for the top-left cell of the first of them, under the
 * names the sheets have now. A live format is that already and is handed back as it is; a
 * format of the file is read through `space`. Undefined for a format none of whose cells is left.
 */
export function conditionalFormatNow(format: ConditionalFormat, space: RuleSpace): ConditionalFormat | undefined {
  if (!format || !Array.isArray(format.ranges) || !Array.isArray(format.rules)) return undefined;
  if (format.live) return format.ranges.length > 0 ? format : undefined;
  // Formulas are carried through the edits rule by rule, the ranges once for all rules.
  const counts: number[] = [];
  const texts: string[] = [];
  for (const rule of format.rules) {
    const formulas = rule && Array.isArray(rule.formulas) ? rule.formulas.map(String) : [];
    counts.push(formulas.length);
    for (const text of formulas) texts.push(text);
  }
  const ranges = movedThroughSpace(format.ranges, texts, space);
  if (!ranges) return undefined;
  const untouched = space.own.length === 0 && space.foreign.length === 0 && !space.nameMap;
  let at = 0;
  const rules = format.rules.map((rule, i) => {
    const formulas = texts.slice(at, at + counts[i]);
    at += counts[i];
    if (!rule) return rule;
    const moved = untouched ? rule : withThresholds(rule, (values) => thresholdsNow(values, space));
    return { ...moved, formulas };
  });
  return { ...format, ranges, rules };
}

// ---------------------------------------------------------------------------
// Data validation

/** Most rules one sheet holds: as many as are read from a file, which is what the format allows. */
export const MAX_SHEET_VALIDATIONS = 65534;
/** The limits Excel has for the messages of a rule and for a list written into it. */
export const MAX_VALIDATION_TITLE = 32;
export const MAX_VALIDATION_MESSAGE = 255;
export const MAX_VALIDATION_LIST = 255;

export const VALIDATION_TYPES = ['none', 'whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom'] as const;
export const VALIDATION_OPERATORS = ['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'] as const;
export type ValidationType = (typeof VALIDATION_TYPES)[number];
export type ValidationOperator = (typeof VALIDATION_OPERATORS)[number];

/**
 * What setValidation() is told about a rule. Formulas are as files store them: English
 * function names, "," between arguments, "." in numbers, with or without a leading "=". A
 * list is either a formula that gives cells ("$A$1:$A$5", a name) or its items in one quoted
 * text ('"yes,no,maybe"'). Dates and times are their serial numbers.
 */
export interface ValidationSettings {
  type: ValidationType;
  /** For numbers, dates, times and text lengths. Default "between". */
  operator?: ValidationOperator;
  formula1?: string;
  /** The upper limit of "between" and "notBetween". */
  formula2?: string;
  /** An empty cell is fine whatever the rule says. Default true. */
  allowBlank?: boolean;
  /** A list shows no arrow in the cell. Default false. */
  hideDropDown?: boolean;
  /** Default true; without a title and a text nothing is shown either way. */
  showInputMessage?: boolean;
  promptTitle?: string;
  prompt?: string;
  /** Default true: a value that breaks the rule is refused (or, by `errorStyle`, only reported). */
  showErrorMessage?: boolean;
  errorStyle?: 'stop' | 'warning' | 'information';
  errorTitle?: string;
  error?: string;
}

const bareFormula = (text: unknown): string | undefined => {
  if (typeof text !== 'string') return undefined;
  const trimmed = text.trim();
  const bare = trimmed.charAt(0) === '=' ? trimmed.slice(1).trim() : trimmed;
  return bare === '' ? undefined : bare;
};

/** A formula of a rule that can be read; throws RangeError for one that cannot. */
function checkedFormula(text: string, what: string): string {
  try {
    parseFormula(text);
  } catch (error) {
    throw new RangeError(`The ${what} of the rule cannot be read as a formula`);
  }
  return text;
}

/** The inner text of a formula that is one quoted text and nothing else; undefined for any other formula. */
function quotedText(formula: string): string | undefined {
  if (formula.length < 2 || formula.charAt(0) !== '"' || formula.charAt(formula.length - 1) !== '"') return undefined;
  const inner = formula.slice(1, -1);
  return /(^|[^"])"([^"]|$)/.test(inner.replace(/""/g, '')) ? undefined : inner.replace(/""/g, '"');
}

/** The rule a command was told, complete and checked; undefined for "no rule". Throws RangeError for what a rule cannot say. */
function checkedValidation(settings: ValidationSettings | undefined): Omit<DataValidation, 'ranges'> | undefined {
  if (settings === undefined || settings === null) return undefined;
  if (typeof settings !== 'object' || VALIDATION_TYPES.indexOf(settings.type) < 0) throw new RangeError('This kind of rule is not known');
  const type = settings.type;
  const message = (value: unknown, limit: number, what: string): string | undefined => {
    if (value === undefined || value === null || value === '') return undefined;
    const text = String(value);
    if (text.length > limit) throw new RangeError(`The ${what} of a rule cannot be longer than ${limit} characters`);
    return text;
  };
  const rule: Omit<DataValidation, 'ranges'> = {
    type,
    operator: 'between',
    allowBlank: settings.allowBlank !== false,
    hideDropDown: type === 'list' && settings.hideDropDown === true,
    showInputMessage: settings.showInputMessage !== false,
    showErrorMessage: settings.showErrorMessage !== false,
    errorStyle: settings.errorStyle === 'warning' || settings.errorStyle === 'information' ? settings.errorStyle : 'stop',
  };
  const promptTitle = message(settings.promptTitle, MAX_VALIDATION_TITLE, 'title of the input message');
  const prompt = message(settings.prompt, MAX_VALIDATION_MESSAGE, 'input message');
  const errorTitle = message(settings.errorTitle, MAX_VALIDATION_TITLE, 'title of the error message');
  const error = message(settings.error, MAX_VALIDATION_MESSAGE, 'error message');
  if (promptTitle !== undefined) rule.promptTitle = promptTitle;
  if (prompt !== undefined) rule.prompt = prompt;
  if (errorTitle !== undefined) rule.errorTitle = errorTitle;
  if (error !== undefined) rule.error = error;

  if (type === 'none') {
    // "Any value" restricts nothing. With a message to show it is still a rule; without one it is none.
    return rule.showInputMessage && (promptTitle !== undefined || prompt !== undefined) ? rule : undefined;
  }
  const first = bareFormula(settings.formula1);
  const second = bareFormula(settings.formula2);
  if (type === 'list') {
    if (first === undefined) throw new RangeError('A list needs a value: its items, or the cells that hold them');
    const items = quotedText(first);
    if (items !== undefined && items.length > MAX_VALIDATION_LIST) throw new RangeError(`The items of a list written into the rule cannot be longer than ${MAX_VALIDATION_LIST} characters together`);
    rule.formula1 = checkedFormula(first, 'source');
    return rule;
  }
  if (type === 'custom') {
    if (first === undefined) throw new RangeError('A custom rule needs a value: its formula');
    rule.formula1 = checkedFormula(first, 'formula');
    return rule;
  }
  const operator = settings.operator === undefined ? 'between' : settings.operator;
  if (VALIDATION_OPERATORS.indexOf(operator) < 0) throw new RangeError('This operator is not known');
  rule.operator = operator;
  const two = operator === 'between' || operator === 'notBetween';
  if (first === undefined || (two && second === undefined)) throw new RangeError(two ? 'The rule needs two values: the lowest and the highest' : 'The rule needs a value');
  rule.formula1 = checkedFormula(first, two ? 'lowest value' : 'value');
  if (two) rule.formula2 = checkedFormula(second as string, 'highest value');
  return rule;
}

const SETTING_KEYS = ['type', 'operator', 'formula1', 'formula2', 'allowBlank', 'hideDropDown', 'showInputMessage', 'promptTitle', 'prompt', 'showErrorMessage', 'errorStyle', 'errorTitle', 'error'] as const;

const sameSettings = (a: Omit<DataValidation, 'ranges'>, b: Omit<DataValidation, 'ranges'>) => SETTING_KEYS.every((key) => a[key] === b[key]);

function worksheetAt(workbook: Workbook, sheetIndex: number, what: string): Sheet {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) throw new RangeError(`No sheet at index ${sheetIndex}`);
  if (sheet.kind !== 'worksheet') throw new RangeError(`Cannot change the ${what} of a ${sheet.kind}`);
  return sheet;
}

/**
 * The formulas of a rule of the file whose ranges are about to change, when they can be kept
 * up to date from here on; throws RangeError for a formula that cannot be read to its end
 * (it could not follow a sheet that gets another name).
 */
function checkLiveFormulas(texts: (string | undefined)[], what: string): void {
  for (const text of texts) {
    if (typeof text === 'string' && text.indexOf('!') >= 0 && formulaNamesSheetsUnread(text)) throw new RangeError(`Cannot change the ${what}: a formula of the rule that is there cannot be read`);
  }
}

/**
 * The data validations that apply to cells of the given ranges now, each as it applies now
 * (see dataValidationNow()), in the order of the sheet. Rules of the extension list are among
 * them (`extended`).
 */
export function validationsIn(workbook: Workbook, sheetIndex: number, ranges: CellRange | readonly CellRange[]): DataValidation[] {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) throw new RangeError(`No sheet at index ${sheetIndex}`);
  const target = targetRanges(ranges);
  const space = ruleSpace(workbook, sheet);
  const out: DataValidation[] = [];
  for (const rule of sheet.dataValidations || []) {
    const now = dataValidationNow(rule, space);
    if (now && rangesIntersect(now.ranges, target)) out.push(now);
  }
  return out;
}

/** Why the data validation of the given ranges cannot be changed, as a message; undefined when it can. */
export function validationProblem(workbook: Workbook, sheetIndex: number, ranges: CellRange | readonly CellRange[]): string | undefined {
  const sheet = worksheetAt(workbook, sheetIndex, 'data validation');
  const target = targetRanges(ranges);
  const space = ruleSpace(workbook, sheet);
  for (const rule of sheet.dataValidations || []) {
    if (!rule || !rule.extended) continue;
    const now = dataValidationNow(rule, space);
    if (now && rangesIntersect(now.ranges, target)) {
      return `Cannot change the data validation of ${now.ranges.map(rangeName).join(' ')}: the rule there is one of the extension list (its source is on another sheet), which is kept as it is`;
    }
  }
  return undefined;
}

/**
 * Sets the data validation of ranges of a sheet: afterwards `settings` is the rule of every
 * cell of them, and of no other cell. As in Excel, the cells are first taken out of the rules
 * they were in: a rule that keeps other cells stays for those (its formulas written for its
 * new first cell), a rule left without cells is gone. Without settings, and for "any value"
 * with no input message to show, the cells are left without a rule.
 *
 * Returns no operation when nothing would change. Throws RangeError for a rule that cannot be
 * stated (an unknown kind or operator, a missing value, a formula that cannot be read, a
 * message that is too long), for ranges outside the sheet, for cells under a rule of the
 * extension list, and for a sheet that has no room for another rule.
 */
export function setValidation(workbook: Workbook, sheetIndex: number, ranges: CellRange | readonly CellRange[], settings: ValidationSettings | undefined): Op[] {
  const sheet = worksheetAt(workbook, sheetIndex, 'data validation');
  const target = targetRanges(ranges);
  const wanted = checkedValidation(settings);
  const problem = validationProblem(workbook, sheetIndex, target);
  if (problem) throw new RangeError(problem);

  const space = ruleSpace(workbook, sheet);
  const current = sheet.dataValidations || [];
  // The rules the cells are in now.
  const touched: { at: number; now: DataValidation }[] = [];
  current.forEach((rule, at) => {
    const now = dataValidationNow(rule, space);
    if (now && rangesIntersect(now.ranges, target)) touched.push({ at, now });
  });
  if (touched.length === 0 && !wanted) return [];
  // One rule that has exactly these cells is the rule that is edited: it stays the rule it is - its place in the
  // order, the element it was read from with what that holds beside the fields of the model - and says something else.
  const same = touched.length === 1 && sameRanges(touched[0].now.ranges, target) ? touched[0] : undefined;
  if (same && wanted) {
    if (sameSettings(same.now, wanted)) return [];
    const edited: DataValidation = { ...same.now, live: true };
    for (const key of ['formula1', 'formula2', 'promptTitle', 'prompt', 'errorTitle', 'error'] as const) delete edited[key];
    Object.assign(edited, wanted);
    checkLiveFormulas([edited.formula1, edited.formula2], 'data validation');
    const list = current.slice();
    list[same.at] = edited;
    return [{ type: 'setDataValidations', sheet: sheetIndex, validations: list }];
  }
  const cut = new Map<number, DataValidation | undefined>();
  for (const { at, now } of touched) {
    const left = subtractRanges(now.ranges, target);
    if (left.length === 0) {
      cut.set(at, undefined);
      continue;
    }
    checkLiveFormulas([now.formula1, now.formula2], 'data validation');
    const kept: DataValidation = { ...now, ranges: left, live: true };
    // The formulas are written for the first cell of the rule, which may be another one now.
    const dRow = left[0].top - now.ranges[0].top;
    const dCol = left[0].left - now.ranges[0].left;
    if (dRow !== 0 || dCol !== 0) {
      for (const name of ['formula1', 'formula2'] as const) if (typeof kept[name] === 'string') kept[name] = shiftFormula(kept[name] as string, dRow, dCol, true);
    }
    cut.set(at, kept);
  }
  const next: DataValidation[] = [];
  current.forEach((rule, at) => {
    // A rule whose cells went with their rows stays in the list as it is: undo brings the cells back.
    if (!cut.has(at)) next.push(rule);
    else if (cut.get(at)) next.push(cut.get(at) as DataValidation);
  });
  if (wanted) {
    // Behind the rules of the main part and before those of the extension list, which is where a file has it.
    let at = next.length;
    while (at > 0 && next[at - 1].extended) at--;
    next.splice(at, 0, { ...wanted, ranges: target, live: true });
  }
  if (next.length > MAX_SHEET_VALIDATIONS) throw new RangeError('The sheet has too many data validations');
  return [{ type: 'setDataValidations', sheet: sheetIndex, validations: next.length > 0 ? next : undefined }];
}

// ---------------------------------------------------------------------------
// Live rules and the operations that move cells or name sheets

/** The lists of rules of one sheet, as an operation replaces them or puts them back. A list that is not stated stays. */
export interface RuleLists {
  sheet: Sheet;
  validations?: { list: DataValidation[] | undefined };
  conditional?: { list: ConditionalFormat[] | undefined };
}

/** A list with some of its entries replaced (undefined: taken out); the list itself when no entry changed. */
function mapList<T>(list: T[] | undefined, change: (item: T) => T | undefined): T[] | undefined {
  if (!list) return list;
  let out: T[] | undefined;
  for (let i = 0; i < list.length; i++) {
    const next = change(list[i]);
    if (next === list[i] && !out) continue;
    if (!out) out = list.slice(0, i);
    if (next !== undefined) out.push(next);
  }
  if (!out) return list;
  return out.length > 0 ? out : undefined;
}

/**
 * What rows or columns inserted into or deleted from `edited` do to the live rules of the
 * workbook: those of the sheet itself move with their cells (a rule left without cells is
 * gone), and the formulas of all of them follow the cells they name. `edit` names the sheet
 * as it is called now. Nothing is changed here: the lists that come out different are
 * returned, for swapRuleLists().
 */
export function liveRulesMoved(workbook: Workbook, edited: Sheet, edit: StructuralEdit): RuleLists[] {
  const out: RuleLists[] = [];
  for (const sheet of workbook.sheets) {
    const own = sheet === edited;
    const validations = mapList(sheet.dataValidations, (rule) => {
      if (!rule || !rule.live || !Array.isArray(rule.ranges) || rule.ranges.length === 0) return rule;
      let ranges = rule.ranges;
      if (own) {
        const moved: CellRange[] = [];
        let same = true;
        for (const range of rule.ranges) {
          const next = adjustRange(range, edit);
          if (!next) {
            same = false;
            continue;
          }
          if (next.top !== range.top || next.left !== range.left || next.bottom !== range.bottom || next.right !== range.right) same = false;
          moved.push(next);
        }
        if (moved.length === 0) return undefined;
        if (!same) ranges = moved;
      }
      const before = { row: rule.ranges[0].top, col: rule.ranges[0].left };
      const after = { row: ranges[0].top, col: ranges[0].left };
      let next = ranges === rule.ranges ? rule : { ...rule, ranges };
      for (const name of ['formula1', 'formula2'] as const) {
        const text = rule[name];
        if (typeof text !== 'string') continue;
        const moved = adjustAnchoredFormula(text, sheet.name, edit, before, after, true);
        if (moved === text) continue;
        if (next === rule) next = { ...rule };
        next[name] = moved;
      }
      return next;
    });
    const conditional = mapList(sheet.conditionalFormats, (format) => {
      if (!format || !format.live || !Array.isArray(format.ranges) || format.ranges.length === 0) return format;
      let ranges = format.ranges;
      if (own) {
        const moved: CellRange[] = [];
        let same = true;
        for (const range of format.ranges) {
          const next = adjustRange(range, edit);
          if (!next) {
            same = false;
            continue;
          }
          if (next.top !== range.top || next.left !== range.left || next.bottom !== range.bottom || next.right !== range.right) same = false;
          moved.push(next);
        }
        if (moved.length === 0) return undefined;
        if (!same) ranges = moved;
      }
      const before = { row: format.ranges[0].top, col: format.ranges[0].left };
      const after = { row: ranges[0].top, col: ranges[0].left };
      const rules = mapList(format.rules, (rule) => {
        if (!rule) return rule;
        let next = rule;
        const formulas = rule.formulas.map((text) => adjustAnchoredFormula(text, sheet.name, edit, before, after, true));
        if (formulas.some((text, i) => text !== rule.formulas[i])) next = { ...next, formulas };
        // A threshold stands on its own: it is no cell's neighbour.
        return withThresholds(next, (values) =>
          values.map((value) => {
            if (!value || typeof value.value !== 'string' || textToNumber(value.value) !== undefined) return value;
            const text = adjustFormula(value.value, sheet.name, edit);
            return text === value.value ? value : { ...value, value: text };
          }),
        );
      }) as ConditionalRule[];
      return ranges === format.ranges && rules === format.rules ? format : { ...format, ranges, rules };
    });
    const entry: RuleLists = { sheet };
    if (validations !== sheet.dataValidations) entry.validations = { list: validations };
    if (conditional !== sheet.conditionalFormats) entry.conditional = { list: conditional };
    if (entry.validations || entry.conditional) out.push(entry);
  }
  return out;
}

/**
 * What other names for sheets (see renameSheetsInFormula()) do to the formulas of the live
 * rules of the workbook. `skip` is a sheet whose rules do not matter (it is being deleted).
 * Nothing is changed here: the lists that come out different are returned, for swapRuleLists().
 */
export function liveRulesRenamed(workbook: Workbook, map: SheetNameMap, skip?: Sheet): RuleLists[] {
  const out: RuleLists[] = [];
  const renamed = (text: string) => (text.indexOf('!') < 0 ? text : renameSheetsInFormula(text, map));
  for (const sheet of workbook.sheets) {
    if (sheet === skip) continue;
    const validations = mapList(sheet.dataValidations, (rule) => {
      if (!rule || !rule.live) return rule;
      let next = rule;
      for (const name of ['formula1', 'formula2'] as const) {
        const text = rule[name];
        if (typeof text !== 'string') continue;
        const now = renamed(text);
        if (now === text) continue;
        if (next === rule) next = { ...rule };
        next[name] = now;
      }
      return next;
    });
    const conditional = mapList(sheet.conditionalFormats, (format) => {
      if (!format || !format.live) return format;
      const rules = mapList(format.rules, (rule) => {
        if (!rule) return rule;
        let next = rule;
        const formulas = rule.formulas.map(renamed);
        if (formulas.some((text, i) => text !== rule.formulas[i])) next = { ...next, formulas };
        return withThresholds(next, (values) =>
          values.map((value) => {
            if (!value || typeof value.value !== 'string') return value;
            const text = renamed(value.value);
            return text === value.value ? value : { ...value, value: text };
          }),
        );
      }) as ConditionalRule[];
      return rules === format.rules ? format : { ...format, rules };
    });
    const entry: RuleLists = { sheet };
    if (validations !== sheet.dataValidations) entry.validations = { list: validations };
    if (conditional !== sheet.conditionalFormats) entry.conditional = { list: conditional };
    if (entry.validations || entry.conditional) out.push(entry);
  }
  return out;
}

/** Puts lists of rules in place and returns the lists that were there, last first: what puts them back. */
export function swapRuleLists(changes: readonly RuleLists[] | undefined): RuleLists[] {
  const undo: RuleLists[] = [];
  for (const change of changes || []) {
    const sheet = change.sheet;
    const previous: RuleLists = { sheet };
    if (change.validations) {
      previous.validations = { list: sheet.dataValidations };
      if (change.validations.list && change.validations.list.length > 0) sheet.dataValidations = change.validations.list;
      else delete sheet.dataValidations;
    }
    if (change.conditional) {
      previous.conditional = { list: sheet.conditionalFormats };
      if (change.conditional.list && change.conditional.list.length > 0) sheet.conditionalFormats = change.conditional.list;
      else delete sheet.conditionalFormats;
    }
    undo.push(previous);
  }
  return undo.reverse();
}

// ---------------------------------------------------------------------------
// Conditional formats

/** Most rules one sheet holds. */
export const MAX_SHEET_CONDITIONAL_RULES = 65534;
/** Longest text a rule looks for in a cell. */
export const MAX_CONDITIONAL_TEXT = 255;

/** The kinds of rule that are made and changed here. Icon sets and time periods are kept, listed, moved and deleted. */
export const CONDITIONAL_TYPES = [
  'cellIs',
  'containsText',
  'notContainsText',
  'beginsWith',
  'endsWith',
  'duplicateValues',
  'uniqueValues',
  'top10',
  'aboveAverage',
  'containsBlanks',
  'notContainsBlanks',
  'containsErrors',
  'notContainsErrors',
  'expression',
  'colorScale',
  'dataBar',
] as const;
export type ConditionalType = (typeof CONDITIONAL_TYPES)[number];

/** What a rule does to the cells it matches: colours as "#RRGGBB". Only what is stated is applied on top of the format of the cell. */
export interface RuleFormat {
  /** The colour of the text. */
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** The colour the cell is filled with. */
  fill?: string;
}

/**
 * What addConditionalRule() and updateConditionalRule() are told about a rule. Formulas are
 * as files store them (English function names, "," between arguments, "." in numbers), with
 * or without a leading "=", written for the top-left cell of the first range.
 */
export interface ConditionalSettings {
  type: ConditionalType;
  /** cellIs: one of the eight comparisons. Default "between". */
  operator?: ValidationOperator;
  /** cellIs: the value, or the lowest and the highest; expression: the formula. */
  formulas?: string[];
  /** containsText, notContainsText, beginsWith, endsWith: the text looked for. */
  text?: string;
  /** top10: how many (default 10), whether that is a percentage, and whether the lowest are meant. */
  rank?: number;
  percent?: boolean;
  bottom?: boolean;
  /** aboveAverage: false for "below average"; `equalAverage` counts the average itself. */
  aboveAverage?: boolean;
  equalAverage?: boolean;
  /** Rules of lower priority are not applied to a cell this rule matches. */
  stopIfTrue?: boolean;
  /**
   * The format the rule gives. For a rule that is changed, only what is stated differently from
   * what the rule has is changed (see updateConditionalRule()); left out, the rule keeps the
   * format it has. `null`: no format at all.
   */
  format?: RuleFormat | null;
  /** colorScale: two or three colours, from the lowest value to the highest; dataBar: the colour of the bar. */
  colors?: string[];
}

/** One rule of a sheet, for whoever lists them. */
export interface ConditionalRuleInfo {
  /** The place of the rule in the order the rules are applied in, 0 first: what the commands name it by. */
  index: number;
  /** The cells the rule applies to now. */
  ranges: CellRange[];
  /** The rule as it applies now: its formulas written for the first cell of `ranges`. */
  rule: ConditionalRule;
  /** The format the rule gives, when it gives one. */
  format: DifferentialFormat | undefined;
  /** What it says and where it applies can be changed. */
  editable: boolean;
  removable: boolean;
  /** Its place in the order can be changed. */
  movable: boolean;
}

interface Slot {
  /** Index of the format in Sheet.conditionalFormats, and of the rule in it. */
  f: number;
  r: number;
  /** The format and the rule as they apply now. */
  ranges: CellRange[];
  rule: ConditionalRule;
}

const priorityOf = (rule: ConditionalRule) => (rule.priority >= 0 ? rule.priority : Infinity);

/** The rules of a sheet that apply to cells now, in the order they are applied in. */
function slotsInOrder(workbook: Workbook, sheet: Sheet): Slot[] {
  const space = ruleSpace(workbook, sheet);
  const slots: Slot[] = [];
  (sheet.conditionalFormats || []).forEach((format, f) => {
    const now = conditionalFormatNow(format, space);
    if (!now) return;
    now.rules.forEach((rule, r) => {
      if (rule) slots.push({ f, r, ranges: now.ranges, rule });
    });
  });
  // Lower numbers first; rules without one keep the order of the file, after the others (as they are evaluated).
  return slots
    .map((slot, order) => ({ slot, order }))
    .sort((a, b) => {
      const pa = priorityOf(a.slot.rule);
      const pb = priorityOf(b.slot.rule);
      return pa === pb ? a.order - b.order : pa < pb ? -1 : 1;
    })
    .map((entry) => entry.slot);
}

const isMade = (type: string): type is ConditionalType => (CONDITIONAL_TYPES as readonly string[]).indexOf(type) >= 0;
/** A rule the extension list has something to say about: there alone (`extended`), or beside the main part (`extId`). */
const inExtension = (rule: ConditionalRule) => rule.extended === true || rule.extId !== undefined;
/**
 * Whether the priorities of a sheet can be given anew: not while a rule whose number is in the
 * extension list applies to cells. (One whose cells are gone is not written either.)
 */
const canRenumber = (slots: Slot[]) => !slots.some((slot) => slot.rule.extended === true);

/**
 * The conditional rules of a sheet in the order they are applied in, each as it applies now;
 * with ranges, only the rules that apply to a cell of them. `index` is what the commands name
 * a rule by; it is the same with and without ranges.
 */
export function conditionalRules(workbook: Workbook, sheetIndex: number, ranges?: CellRange | readonly CellRange[]): ConditionalRuleInfo[] {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) throw new RangeError(`No sheet at index ${sheetIndex}`);
  const target = ranges === undefined ? undefined : targetRanges(ranges);
  const slots = slotsInOrder(workbook, sheet);
  const renumber = canRenumber(slots);
  const dxfs = workbook.styles.dxfFormats;
  const out: ConditionalRuleInfo[] = [];
  slots.forEach((slot, index) => {
    if (target && !rangesIntersect(slot.ranges, target)) return;
    const rule = slot.rule;
    const kept = inExtension(rule);
    // A rule only the extension list has is not touched at all; one it says more about can go, and go elsewhere, with its counterpart there.
    const fixed = rule.extended === true;
    out.push({
      index,
      ranges: slot.ranges,
      rule,
      format: rule.dxf || (rule.dxfId !== undefined && dxfs ? dxfs[rule.dxfId] : undefined),
      editable: !kept && isMade(rule.type),
      removable: !fixed,
      movable: renumber && rule.extended !== true,
    });
  });
  return out;
}

const HEX_COLOR = /^#([0-9a-fA-F]{6})$/;

function colorOf(text: unknown, what: string): Color {
  const match = typeof text === 'string' ? HEX_COLOR.exec(text) : null;
  if (!match) throw new RangeError(`The ${what} is written #RRGGBB`);
  return { rgb: `FF${match[1].toUpperCase()}` };
}

/** A colour of the file as a dialog shows it ("#RRGGBB"); undefined for one it cannot show. */
export type ColorReader = (color: Color | undefined) => string | undefined;

const statedColor: ColorReader = (color) => (color && typeof color.rgb === 'string' && /^[0-9A-Fa-f]{6,8}$/.test(color.rgb) ? `#${color.rgb.slice(-6).toUpperCase()}` : undefined);

/**
 * What a format says in the words of RuleFormat: the part of it that the dialogs show and the
 * commands change. What it says beside that (a border, a number format, a gradient, a font
 * size) is not in it, and neither is a property that is stated as off.
 */
export function ruleFormatOf(format: DifferentialFormat | undefined, read: ColorReader = statedColor): RuleFormat {
  const out: RuleFormat = {};
  if (!format) return out;
  const font = format.font;
  if (font) {
    const color = read(font.color) || statedColor(font.color);
    if (color !== undefined) out.color = color;
    if (font.bold === true) out.bold = true;
    if (font.italic === true) out.italic = true;
    if (font.strike === true) out.strike = true;
    if (font.underline !== undefined && font.underline !== 'none') out.underline = true;
  }
  const fill = format.fill;
  if (fill && fill.type === 'pattern' && fill.pattern !== 'none') {
    const color = read(fill.bg) || statedColor(fill.bg) || read(fill.fg) || statedColor(fill.fg);
    if (color !== undefined) out.fill = color;
  }
  return out;
}

/** Plain data with the keys of every object in order: two values that say the same give the same text. */
function stable(value: unknown): string {
  return JSON.stringify(value, (_key, item) => {
    if (item && typeof item === 'object' && !Array.isArray(item)) {
      const sorted: Record<string, unknown> = {};
      for (const key of Object.keys(item).sort()) if (item[key] !== undefined) sorted[key] = item[key];
      return sorted;
    }
    return item;
  });
}

/**
 * The differential format a rule gets: `base` (the format it has) with what `format` states
 * differently from it. A property that is stated as the rule shows it stays what it is in the
 * file (a colour of the theme, a "bold" that is stated as off); what the words of RuleFormat
 * cannot say (border, number format, gradient) is never touched. Undefined for nothing at all.
 */
function formatOf(format: RuleFormat, base: DifferentialFormat | undefined, read: ColorReader): DifferentialFormat | undefined {
  if (typeof format !== 'object') throw new RangeError('The format of a rule is a set of properties');
  const shown = ruleFormatOf(base, read);
  const out: DifferentialFormat = base ? { ...base } : {};
  const font: NonNullable<DifferentialFormat['font']> = base && base.font ? { ...base.font } : {};
  for (const key of ['bold', 'italic', 'strike'] as const) {
    const wanted = format[key];
    // Not stated, and not shown as on: as it is (which may be "stated as off").
    if (wanted === undefined ? font[key] !== true : wanted === font[key]) continue;
    if (wanted === undefined) delete font[key];
    else font[key] = wanted === true;
  }
  if ((format.underline === true) !== (shown.underline === true) || (format.underline === false && font.underline === undefined)) {
    if (format.underline === undefined) delete font.underline;
    else font.underline = format.underline === true ? 'single' : 'none';
  }
  const same = (a: string | undefined, b: string | undefined) => (a === undefined ? b === undefined : b !== undefined && a.toUpperCase() === b.toUpperCase());
  if (!same(format.color, shown.color)) {
    if (format.color === undefined) delete font.color;
    else font.color = colorOf(format.color, 'colour of the text');
  }
  if (Object.keys(font).length > 0) out.font = font;
  else delete out.font;
  if (!same(format.fill, shown.fill)) {
    if (format.fill === undefined) delete out.fill;
    else out.fill = { type: 'pattern', pattern: 'solid', bg: colorOf(format.fill, 'fill colour') };
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

/** A rule with the format `format` asks for (see ConditionalSettings.format); the rule itself when that is the format it has. */
function withFormat(rule: ConditionalRule, format: RuleFormat | null, styles: StyleSheet, read: ColorReader): ConditionalRule {
  const dxfs = styles.dxfFormats;
  const base = rule.dxf || (rule.dxfId !== undefined && dxfs ? dxfs[rule.dxfId] : undefined);
  if (format === null) {
    if (rule.dxf === undefined && rule.dxfId === undefined) return rule;
    const bare = { ...rule };
    delete bare.dxf;
    delete bare.dxfId;
    return bare;
  }
  const wanted = formatOf(format, base, read);
  if (stable(wanted) === stable(base)) return rule;
  if (!wanted) return withFormat(rule, null, styles, read);
  // The record the format started from is still named: what it holds beside font and fill is taken from it when the file is written.
  return { ...rule, dxf: wanted };
}

const inQuotes = (text: string) => `"${text.replace(/"/g, '""')}"`;

/** The formula Excel writes for a rule that is stated by its kind alone, for the cell `cell` ("A1"). */
function standardFormula(type: ConditionalType, text: string, cell: string): string | undefined {
  switch (type) {
    case 'containsText':
      return `NOT(ISERROR(SEARCH(${inQuotes(text)},${cell})))`;
    case 'notContainsText':
      return `ISERROR(SEARCH(${inQuotes(text)},${cell}))`;
    case 'beginsWith':
      return `LEFT(${cell},LEN(${inQuotes(text)}))=${inQuotes(text)}`;
    case 'endsWith':
      return `RIGHT(${cell},LEN(${inQuotes(text)}))=${inQuotes(text)}`;
    case 'containsBlanks':
      return `LEN(TRIM(${cell}))=0`;
    case 'notContainsBlanks':
      return `LEN(TRIM(${cell}))>0`;
    case 'containsErrors':
      return `ISERROR(${cell})`;
    case 'notContainsErrors':
      return `NOT(ISERROR(${cell}))`;
    default:
      return undefined;
  }
}

/** The word the file has for the comparison of a text rule (ST_ConditionalFormattingOperator). */
const TEXT_OPERATORS: Record<string, string> = { containsText: 'containsText', notContainsText: 'notContains', beginsWith: 'beginsWith', endsWith: 'endsWith' };

const cellText = (row: number, col: number) => {
  let name = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return `${name}${row + 1}`;
};

/**
 * The cell a text rule's formula is written for, as the formula spells it ("A1", "$A1", ...):
 * the formula is then the one Excel writes for the text of the rule, and the rule can be
 * stated by its text alone. Undefined for any other formula. A rule without a formula is
 * answered from its text: the plain name of the cell.
 */
export function textRuleCell(rule: Pick<ConditionalRule, 'type' | 'text' | 'formulas'>, anchor: { row: number; col: number }): string | undefined {
  if (!isMade(rule.type) || TEXT_OPERATORS[rule.type] === undefined || typeof rule.text !== 'string') return undefined;
  const written = typeof rule.formulas[0] === 'string' ? rule.formulas[0].trim() : undefined;
  const plain = cellText(anchor.row, anchor.col);
  if (written === undefined) return plain;
  const letters = plain.replace(/\d+$/, '');
  const digits = plain.slice(letters.length);
  for (const cell of [plain, `$${letters}${digits}`, `${letters}$${digits}`, `$${letters}$${digits}`]) {
    if (standardFormula(rule.type, rule.text, cell) === written) return cell;
  }
  return undefined;
}

/**
 * The rule the settings stand for, for a rule whose first cell is `anchor`. `base` is the rule
 * that is changed. While the kind of rule stays, the settings are a patch on it: what they do
 * not state, and what they state as the rule has it, stays as the rule has it - with everything
 * the settings have no word for (the thresholds of a scale or a bar, the lengths of a bar, a
 * number of deviations, the spelling of a formula). Its place in the file and the record of
 * its format stay in any case. Throws RangeError for what a rule cannot say.
 */
function ruleFromSettings(settings: ConditionalSettings, anchor: { row: number; col: number }, base: ConditionalRule | undefined, styles: StyleSheet, read: ColorReader): ConditionalRule {
  if (!settings || typeof settings !== 'object' || !isMade(settings.type)) throw new RangeError('This kind of rule is not known, or is not made here');
  const type = settings.type;
  const same = base !== undefined && base.type === type ? base : undefined;
  let rule: ConditionalRule;
  if (same) rule = { ...same, formulas: same.formulas.slice() };
  else {
    rule = { type, priority: base ? base.priority : Infinity, stopIfTrue: base ? base.stopIfTrue : false, formulas: [], percent: false, bottom: false, aboveAverage: true, equalAverage: false };
    if (base && base.src) rule.src = base.src;
    if (base && base.dxfId !== undefined) rule.dxfId = base.dxfId;
    if (base && base.dxf) rule.dxf = base.dxf;
  }
  if (settings.stopIfTrue !== undefined) rule.stopIfTrue = settings.stopIfTrue === true;
  const cell = cellText(anchor.row, anchor.col);
  const given = Array.isArray(settings.formulas) ? settings.formulas.map(bareFormula) : [];
  /** A formula as it is given; one that is the formula the rule has is that formula, and one that is not given is too. */
  const formula = (at: number, what: string): string | undefined => {
    const kept = same ? same.formulas[at] : undefined;
    const text = given[at] !== undefined ? given[at] : kept;
    return text === undefined || text === kept ? text : checkedFormula(text, what);
  };
  switch (type) {
    case 'cellIs': {
      const operator = settings.operator !== undefined ? settings.operator : same && typeof same.operator === 'string' ? (same.operator as ValidationOperator) : 'between';
      if (VALIDATION_OPERATORS.indexOf(operator) < 0) throw new RangeError('This operator is not known');
      const two = operator === 'between' || operator === 'notBetween';
      const first = formula(0, two ? 'lowest value' : 'value');
      const second = two ? formula(1, 'highest value') : undefined;
      if (first === undefined || (two && second === undefined)) throw new RangeError(two ? 'The rule needs two values: the lowest and the highest' : 'The rule needs a value');
      rule.operator = operator;
      rule.formulas = two ? [first, second as string] : [first];
      break;
    }
    case 'expression': {
      const first = formula(0, 'formula');
      if (first === undefined) throw new RangeError('The rule needs a value: its formula');
      rule.formulas = [first];
      break;
    }
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith': {
      const text = typeof settings.text === 'string' ? settings.text : same && typeof same.text === 'string' ? same.text : '';
      if (text === '') throw new RangeError('The rule needs a text to look for');
      if (text.length > MAX_CONDITIONAL_TEXT) throw new RangeError(`The text of a rule cannot be longer than ${MAX_CONDITIONAL_TEXT} characters`);
      // The text the rule has: the rule as it is, with the formula as it is written.
      if (same && same.text === text) break;
      rule.text = text;
      rule.operator = TEXT_OPERATORS[type];
      // Written for the cell the old formula named, the way it named it ("$A1").
      rule.formulas = [standardFormula(type, text, (same && textRuleCell(same, anchor)) || cell) as string];
      break;
    }
    case 'containsBlanks':
    case 'notContainsBlanks':
    case 'containsErrors':
    case 'notContainsErrors':
      if (!same) rule.formulas = [standardFormula(type, '', cell) as string];
      break;
    case 'top10': {
      const rank = settings.rank !== undefined ? settings.rank : same && same.rank !== undefined ? same.rank : 10;
      const percent = settings.percent !== undefined ? settings.percent === true : same ? same.percent === true : false;
      const limit = percent ? 100 : 1000;
      if (typeof rank !== 'number' || Math.floor(rank) !== rank || rank < 1 || rank > limit) throw new RangeError(`The number of a top or bottom rule is a whole number between 1 and ${limit}`);
      rule.rank = rank;
      rule.percent = percent;
      rule.bottom = settings.bottom !== undefined ? settings.bottom === true : same ? same.bottom === true : false;
      break;
    }
    case 'aboveAverage':
      if (settings.aboveAverage !== undefined || !same) rule.aboveAverage = settings.aboveAverage !== false;
      if (settings.equalAverage !== undefined || !same) rule.equalAverage = settings.equalAverage === true;
      break;
    case 'colorScale': {
      const kept = same && same.colorScale ? same.colorScale : undefined;
      if (settings.colors === undefined && kept) break;
      const colors = Array.isArray(settings.colors) ? settings.colors : [];
      if (colors.length !== 2 && colors.length !== 3) throw new RangeError('A colour scale has two or three colours');
      const wanted = colors.map((color) => colorOf(color, 'colour of a colour scale'));
      if (kept && kept.colors.length === colors.length && kept.values.length === colors.length) {
        // As many colours as there were: the thresholds stay, and so does every colour that is shown as it is given.
        const next = kept.colors.map((color, i) => ((read(color) || statedColor(color) || '') === colors[i].toUpperCase() ? color : wanted[i]));
        if (next.some((color, i) => color !== kept.colors[i])) rule.colorScale = { ...kept, colors: next };
        break;
      }
      const values: ConditionalValue[] = [{ type: 'min', gte: true }];
      if (colors.length === 3) values.push({ type: 'percentile', value: '50', gte: true });
      values.push({ type: 'max', gte: true });
      rule.colorScale = { values, colors: wanted };
      break;
    }
    case 'dataBar': {
      const kept = same && same.dataBar ? same.dataBar : undefined;
      if (settings.colors === undefined && kept) break;
      const colors = Array.isArray(settings.colors) ? settings.colors : [];
      if (colors.length < 1) throw new RangeError('A data bar needs a colour');
      const color = colorOf(colors[0], 'colour of a data bar');
      if (kept) {
        // The colour alone: where the bar starts and ends, how long it may be and whether the value is shown stay.
        if ((read(kept.color) || statedColor(kept.color) || '') !== colors[0].toUpperCase()) rule.dataBar = { ...kept, color };
        break;
      }
      rule.dataBar = { min: { type: 'min', gte: true }, max: { type: 'max', gte: true }, color, showValue: true, minLength: 10, maxLength: 90, extended: false, gradient: true, negativeFillSameAsPositive: false, negativeBorderSameAsPositive: true, axisPosition: 'automatic', direction: 'context' };
      break;
    }
    default:
    // duplicateValues and uniqueValues say everything by their kind.
  }
  // Scales and bars colour by themselves; every other rule gives the format it is told, or keeps the one it has.
  if (type === 'colorScale' || type === 'dataBar') return withFormat(rule, null, styles, read);
  return settings.format === undefined ? rule : withFormat(rule, settings.format, styles, read);
}

/** The formulas of a rule (its own, and those of its thresholds) that live rules are kept up to date by. */
function ruleFormulaTexts(rule: ConditionalRule): string[] {
  const out = rule.formulas.slice();
  const values = rule.colorScale ? rule.colorScale.values : rule.iconSet ? rule.iconSet.values : rule.dataBar ? [rule.dataBar.min, rule.dataBar.max] : [];
  for (const value of values) if (value && typeof value.value === 'string') out.push(value.value);
  return out;
}

/** A rule with its formulas written for another first cell: relative references keep meaning the cells they meant. */
function reanchored(rule: ConditionalRule, from: CellRange, to: CellRange): ConditionalRule {
  const dRow = to.top - from.top;
  const dCol = to.left - from.left;
  if ((dRow === 0 && dCol === 0) || rule.formulas.length === 0) return rule;
  return { ...rule, formulas: rule.formulas.map((text) => shiftFormula(text, dRow, dCol, true)) };
}

/** What becomes of one rule of the sheet: it stays where it is, is another rule in a format of its own, or goes. */
interface Placed {
  /** Where the rule is in the list now; -1 for a rule that is added. */
  f: number;
  r: number;
  rule: ConditionalRule;
  /** The rule becomes a live format of its own with these ranges. */
  detach?: CellRange[];
}

/**
 * The list of formats once the rules are as `placed` says, in the order they are to be applied
 * in. Rules of the list that `placed` does not name (their cells are gone) stay as they are.
 * `replaced` are formats that changed as a whole (all their rules got other cells): they stand
 * in the place of the ones they come from, and `removed` the places ("format:rule") of rules that
 * are taken out. With `renumber` the rules get the numbers 1, 2, 3... in that order; a rule
 * whose number stays is the object it was.
 */
function rebuiltFormats(sheet: Sheet, placed: Placed[], renumber: boolean, removed: Set<string>, replaced?: Map<number, ConditionalFormat>): ConditionalFormat[] | undefined {
  const current = sheet.conditionalFormats || [];
  const numbered = new Map<Placed, ConditionalRule>();
  placed.forEach((item, i) => {
    const priority = renumber ? i + 1 : item.rule.priority;
    numbered.set(item, item.rule.priority === priority ? item.rule : { ...item.rule, priority });
  });
  const byPlace = new Map<string, Placed>();
  for (const item of placed) if (item.f >= 0) byPlace.set(`${item.f}:${item.r}`, item);
  const out: ConditionalFormat[] = [];
  current.forEach((format, f) => {
    const base = (replaced && replaced.get(f)) || format;
    let same = base === format;
    const rules: ConditionalRule[] = [];
    base.rules.forEach((rule, r) => {
      const key = `${f}:${r}`;
      const item = byPlace.get(key);
      if (!item) {
        // A rule that is taken out goes; one that nobody names is a rule whose cells are gone, which stays (undo brings the cells back).
        if (removed.has(key)) same = false;
        else rules.push(rule);
        return;
      }
      if (item.detach) {
        same = false;
        return;
      }
      const next = numbered.get(item) as ConditionalRule;
      if (next !== rule) same = false;
      rules.push(next);
    });
    if (rules.length === 0) return;
    out.push(same ? format : { ...base, rules });
  });
  for (const item of placed) {
    if (!item.detach) {
      if (item.f < 0) throw new RangeError('A rule that is added needs cells to apply to');
      continue;
    }
    out.push({ ranges: item.detach, rules: [numbered.get(item) as ConditionalRule], live: true });
  }
  let count = 0;
  for (const format of out) count += format.rules.length;
  if (count > MAX_SHEET_CONDITIONAL_RULES) throw new RangeError('The sheet has too many conditional formats');
  return out.length > 0 ? out : undefined;
}

/** Records one style sheet holds for differential formats: their number is written into 16 bits by older readers. */
export const MAX_DIFFERENTIAL_FORMATS = 65536;

function conditionalOp(sheetIndex: number, workbook: Workbook, sheet: Sheet, placed: Placed[], renumber: boolean, removed: Slot[] = [], replaced?: Map<number, ConditionalFormat>): Op[] {
  const places = new Set(removed.map((slot) => `${slot.f}:${slot.r}`));
  const formats = rebuiltFormats(sheet, placed, renumber, places, replaced);
  // Every rule that states its format itself may need a record of its own when the file is written: there is room for so many.
  let own = 0;
  for (const other of workbook.sheets) {
    for (const format of (other === sheet ? formats : other.conditionalFormats) || []) for (const rule of format.rules) if (rule && rule.dxf && !rule.extended) own++;
  }
  if (own > 0 && workbook.styles.dxfs.length + own > MAX_DIFFERENTIAL_FORMATS) throw new RangeError('The workbook has too many formats of conditional rules');
  return [{ type: 'setConditionalFormats', sheet: sheetIndex, formats }];
}

const stay = (slot: Slot, sheet: Sheet): Placed => ({ f: slot.f, r: slot.r, rule: (sheet.conditionalFormats as ConditionalFormat[])[slot.f].rules[slot.r] });

function slotAt(slots: Slot[], index: number): Slot {
  const slot = slots[index];
  if (!slot || Math.floor(index) !== index) throw new RangeError(`No rule at index ${index}`);
  return slot;
}

/** Refuses a change of a rule that only the extension list has; with `says`, of what a rule says that the list says more about. */
function refuseExtension(rule: ConditionalRule, what: string, says = false): void {
  if (rule.extended === true || (says && rule.extId !== undefined)) throw new RangeError(`Cannot ${what} this rule: the extension list has more to say about it, and is kept as it is`);
}

/**
 * Adds a conditional rule for ranges of a sheet. As in Excel the new rule is the first to be
 * applied, and the others move down by one - except in a sheet with a rule only the extension
 * list has, whose numbers stay: there the new rule comes last.
 *
 * Throws RangeError for a rule that cannot be stated, for ranges outside the sheet and for a
 * sheet that has no room for another rule.
 */
export function addConditionalRule(workbook: Workbook, sheetIndex: number, ranges: CellRange | readonly CellRange[], settings: ConditionalSettings): Op[] {
  const sheet = worksheetAt(workbook, sheetIndex, 'conditional formats');
  const target = targetRanges(ranges);
  const rule = ruleFromSettings(settings, { row: target[0].top, col: target[0].left }, undefined, workbook.styles, (color) => resolveColor(color, workbook.styles));
  const slots = slotsInOrder(workbook, sheet);
  const renumber = canRenumber(slots);
  const others = slots.map((slot) => stay(slot, sheet));
  if (renumber) return conditionalOp(sheetIndex, workbook, sheet, [{ f: -1, r: -1, rule, detach: target } as Placed].concat(others), true);
  let highest = 0;
  for (const format of sheet.conditionalFormats || []) for (const other of format.rules) if (other && other.priority < Infinity && other.priority > highest) highest = Math.floor(other.priority);
  return conditionalOp(sheetIndex, workbook, sheet, others.concat([{ f: -1, r: -1, rule: { ...rule, priority: highest + 1 }, detach: target }]), false);
}

/** What updateConditionalRule() changes of a rule; what is not stated stays. */
export interface ConditionalPatch {
  /** The cells the rule applies to. Its formulas are rewritten for the new first cell: a relative reference keeps meaning the same cell for each cell of the rule. */
  ranges?: CellRange | readonly CellRange[];
  stopIfTrue?: boolean;
  /**
   * What the rule says. While the kind of rule stays this is a patch as well: what is left
   * out, or stated as the rule has it, stays as the rule has it, together with everything the
   * settings have no word for (see ConditionalSettings).
   */
  rule?: ConditionalSettings;
  /** The format alone (see ConditionalSettings.format); `rule.format` counts when both are stated. */
  format?: RuleFormat | null;
}

/**
 * Changes the rule at `index` (see conditionalRules()). The rule becomes a format of its own:
 * the other rules of the element it was read from stay in it.
 *
 * Returns no operation when nothing would change: a rule that is stated again the way it is
 * listed is the rule it was. Throws RangeError for an index that names no rule, for a rule only
 * the extension list has, for what a rule says when the extension list says more about it (a
 * data bar of Excel 2010: its cells and `stopIfTrue` can be changed), for a kind of rule that
 * is only kept (likewise), and for what a rule cannot say.
 */
export function updateConditionalRule(workbook: Workbook, sheetIndex: number, index: number, patch: ConditionalPatch): Op[] {
  const sheet = worksheetAt(workbook, sheetIndex, 'conditional formats');
  const slots = slotsInOrder(workbook, sheet);
  const slot = slotAt(slots, index);
  if (!patch || typeof patch !== 'object') return [];
  // Where a rule applies and whether it stops the others is said by the main part; what it says is said by the extension list as well.
  refuseExtension(slot.rule, 'change', patch.rule !== undefined || patch.format !== undefined);
  if ((patch.rule !== undefined || patch.format !== undefined) && !isMade(slot.rule.type)) throw new RangeError(`A rule of the kind "${slot.rule.type}" cannot be changed here: it is kept as it is`);
  const ranges = patch.ranges === undefined ? slot.ranges : targetRanges(patch.ranges);
  const read: ColorReader = (color) => resolveColor(color, workbook.styles);
  let rule = slot.rule;
  // What it says is stated for the cells it has; where it applies comes after.
  if (patch.rule !== undefined) rule = ruleFromSettings(patch.rule, { row: slot.ranges[0].top, col: slot.ranges[0].left }, rule, workbook.styles, read);
  if (patch.format !== undefined && !(patch.rule !== undefined && patch.rule.format !== undefined) && rule.type !== 'colorScale' && rule.type !== 'dataBar') rule = withFormat(rule, patch.format, workbook.styles, read);
  if (patch.stopIfTrue !== undefined && (patch.stopIfTrue === true) !== rule.stopIfTrue) rule = { ...rule, stopIfTrue: patch.stopIfTrue === true };
  const moved = !sameRanges(ranges, slot.ranges);
  if (moved) rule = reanchored(rule, slot.ranges[0], ranges[0]);
  if (!moved && (rule === slot.rule || stable(rule) === stable(slot.rule))) return [];
  checkLiveFormulas(ruleFormulaTexts(rule), 'conditional format');
  const placed = slots.map((other): Placed => (other === slot ? { f: slot.f, r: slot.r, rule, detach: ranges } : stay(other, sheet)));
  return conditionalOp(sheetIndex, workbook, sheet, placed, canRenumber(slots));
}

/** Deletes the rule at `index`. Throws RangeError for an index that names no rule and for a rule only the extension list has. */
export function removeConditionalRule(workbook: Workbook, sheetIndex: number, index: number): Op[] {
  const sheet = worksheetAt(workbook, sheetIndex, 'conditional formats');
  const slots = slotsInOrder(workbook, sheet);
  const slot = slotAt(slots, index);
  refuseExtension(slot.rule, 'delete');
  const placed = slots.filter((other) => other !== slot).map((other) => stay(other, sheet));
  return conditionalOp(sheetIndex, workbook, sheet, placed, canRenumber(slots), [slot]);
}

/**
 * Moves the rule at `index` so that it is applied as number `to` (0: first). Throws RangeError
 * for an index that names no rule, and in a sheet with a rule only the extension list has:
 * the number of that rule is kept as it is, and with it the order of all.
 */
export function moveConditionalRule(workbook: Workbook, sheetIndex: number, index: number, to: number): Op[] {
  const sheet = worksheetAt(workbook, sheetIndex, 'conditional formats');
  const slots = slotsInOrder(workbook, sheet);
  const slot = slotAt(slots, index);
  slotAt(slots, to);
  if (!canRenumber(slots)) throw new RangeError('Cannot change the order of the rules of this sheet: the extension list has a rule of its own, which is kept as it is');
  if (index === to) return [];
  const order = slots.filter((other) => other !== slot);
  order.splice(to, 0, slot);
  return conditionalOp(sheetIndex, workbook, sheet, order.map((other) => stay(other, sheet)), true);
}

/**
 * Takes the cells of the given ranges out of every conditional rule of the sheet, or, without
 * ranges, deletes all its rules. A format that keeps other cells stays for those, with all
 * its rules (their formulas written for its new first cell).
 *
 * Returns no operation when nothing would change. Throws RangeError when a rule only the
 * extension list has is among those that would change.
 */
export function clearConditionalRules(workbook: Workbook, sheetIndex: number, ranges?: CellRange | readonly CellRange[]): Op[] {
  const sheet = worksheetAt(workbook, sheetIndex, 'conditional formats');
  const current = sheet.conditionalFormats;
  if (!current || current.length === 0) return [];
  const target = ranges === undefined ? undefined : targetRanges(ranges);
  const slots = slotsInOrder(workbook, sheet);
  const space = ruleSpace(workbook, sheet);
  const replaced = new Map<number, ConditionalFormat>();
  const gone = new Set<number>();
  current.forEach((format, f) => {
    const now = conditionalFormatNow(format, space);
    if (!now || (target && !rangesIntersect(now.ranges, target))) return;
    for (const rule of now.rules) if (rule) refuseExtension(rule, 'clear');
    const left = target ? subtractRanges(now.ranges, target) : [];
    if (left.length === 0) {
      gone.add(f);
      return;
    }
    for (const rule of now.rules) if (rule) checkLiveFormulas(ruleFormulaTexts(rule), 'conditional format');
    replaced.set(f, { ...now, ranges: left, rules: now.rules.map((rule) => (rule ? reanchored(rule, now.ranges[0], left[0]) : rule)), live: true });
  });
  if (replaced.size === 0 && gone.size === 0) return [];
  const removed = slots.filter((slot) => gone.has(slot.f));
  const placed = slots
    .filter((slot) => !gone.has(slot.f))
    .map((slot): Placed => {
      const format = replaced.get(slot.f);
      return format ? { f: slot.f, r: slot.r, rule: format.rules[slot.r] } : stay(slot, sheet);
    });
  return conditionalOp(sheetIndex, workbook, sheet, placed, canRenumber(slots), removed, replaced);
}

/**
 * A rule as the main part of a sheet states it, without what only the extension list says
 * about it (the border, axis and negative colours of a data bar, icons chosen one by one):
 * what a sheet that is in no file can be written with. Undefined for a rule only the
 * extension list has.
 */
export function mainPartRule(rule: ConditionalRule): ConditionalRule | undefined {
  if (!rule || rule.extended) return undefined;
  const out: ConditionalRule = { ...rule, formulas: rule.formulas.slice() };
  delete out.src;
  delete out.extId;
  if (out.dataBar) {
    const plain = (value: ConditionalValue): ConditionalValue => (value.type === 'autoMin' ? { ...value, type: 'min' } : value.type === 'autoMax' ? { ...value, type: 'max' } : value);
    const bar = out.dataBar;
    out.dataBar = {
      min: plain(bar.min),
      max: plain(bar.max),
      showValue: bar.showValue,
      minLength: bar.minLength,
      maxLength: bar.maxLength,
      extended: false,
      gradient: true,
      negativeFillSameAsPositive: false,
      negativeBorderSameAsPositive: true,
      axisPosition: 'automatic',
      direction: 'context',
      // The colour a bar without one is written with.
      color: bar.color || { rgb: 'FF638EC6' },
    };
  }
  if (out.iconSet && out.iconSet.custom) {
    const icons = { ...out.iconSet };
    delete icons.custom;
    out.iconSet = icons;
  }
  return out;
}

/** How many rules of a sheet the extension list says something about, which a copy of the sheet does not get. */
export function extensionRuleCount(workbook: Workbook, sheet: Sheet): number {
  const space = ruleSpace(workbook, sheet);
  let count = 0;
  for (const format of sheet.conditionalFormats || []) {
    const now = conditionalFormatNow(format, space);
    if (!now) continue;
    for (const rule of now.rules) if (rule && (rule.extended || rule.extId !== undefined || (rule.iconSet && rule.iconSet.custom))) count++;
  }
  return count;
}
