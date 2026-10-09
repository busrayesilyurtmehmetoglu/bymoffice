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
import { adjustAnchoredFormula, adjustFormula, adjustRange, formulaNamesSheetsUnread, formulaReferences, renameSheetsInFormula, shiftFormula, type StructuralEdit } from '../calc/refs';
import { textToNumber } from '../calc/values';
import { ICON_SETS } from '../format/conditional';
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

/** The kinds of rule that are made and changed here. Every other kind of a file is kept, listed, moved and deleted. */
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
  'iconSet',
  'timePeriod',
] as const;
export type ConditionalType = (typeof CONDITIONAL_TYPES)[number];

/**
 * The icon sets a rule is made with: those the main part of a sheet can name (ST_IconSetType).
 * The sets of the 2010 extension ("3Stars", "3Triangles", "5Boxes") and icons chosen one by
 * one are kept where a file has them, and not made here.
 */
export const ICON_SET_NAMES = [
  '3Arrows',
  '3ArrowsGray',
  '3Flags',
  '3TrafficLights1',
  '3TrafficLights2',
  '3Signs',
  '3Symbols',
  '3Symbols2',
  '4Arrows',
  '4ArrowsGray',
  '4RedToBlack',
  '4Rating',
  '4TrafficLights',
  '5Arrows',
  '5ArrowsGray',
  '5Rating',
  '5Quarters',
] as const;
export type IconSetName = (typeof ICON_SET_NAMES)[number];

/** The periods a "dates occurring" rule is made with (ST_TimePeriod). */
export const TIME_PERIODS = ['yesterday', 'today', 'tomorrow', 'last7Days', 'lastWeek', 'thisWeek', 'nextWeek', 'lastMonth', 'thisMonth', 'nextMonth'] as const;
export type TimePeriod = (typeof TIME_PERIODS)[number];

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
  /**
   * iconSet: the set (default "3TrafficLights1"), whether its icons go the other way round, and
   * whether the value is shown beside the icon (default true). A new rule shares the range out
   * evenly, in percent (0/33/67, 0/25/50/75, 0/20/40/60/80); a rule that is changed keeps the
   * thresholds it has while the set has as many icons as before.
   */
  iconSet?: IconSetName;
  reverse?: boolean;
  showValue?: boolean;
  /** timePeriod: the period the date of a cell is in. */
  timePeriod?: TimePeriod;
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
 * The formula Excel writes beside the `timePeriod` of a rule, for the cell `cell` ("A1"). The
 * rule is evaluated by its period (see inTimePeriod() of the conditional formats); the formula
 * is what a reader without the attribute goes by.
 *
 * UNSURE (not compared with Excel): the two part ways for a date that a cell holds as text.
 * The rule looks at numbers only; the formula takes "45000" and "3/15/2023" for the dates
 * they read as. Should Excel go by the formula, it fills such cells and this engine does not.
 *
 * UNSURE (not compared with Excel): `cell` is the first cell of the first range of the rule,
 * as for every formula of a rule here. LibreOffice writes and reads such formulas for the
 * top-left corner of all the ranges together, which is another cell when the first range is
 * not the one that begins there.
 */
function periodFormula(period: string, cell: string): string | undefined {
  switch (period) {
    case 'today':
      return `FLOOR(${cell},1)=TODAY()`;
    case 'yesterday':
      return `FLOOR(${cell},1)=TODAY()-1`;
    case 'tomorrow':
      return `FLOOR(${cell},1)=TODAY()+1`;
    case 'last7Days':
      return `AND(TODAY()-FLOOR(${cell},1)<=6,FLOOR(${cell},1)<=TODAY())`;
    // UNSURE: the formulas of all ten periods are, letter for letter, what other libraries write for them (exceljs all of them,
    // XlsxWriter eight); they were not compared with a file saved by Excel.
    case 'lastWeek':
      return `AND(TODAY()-ROUNDDOWN(${cell},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${cell},0)<(WEEKDAY(TODAY())+7))`;
    case 'thisWeek':
      return `AND(TODAY()-ROUNDDOWN(${cell},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${cell},0)-TODAY()<=7-WEEKDAY(TODAY()))`;
    case 'nextWeek':
      return `AND(ROUNDDOWN(${cell},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${cell},0)-TODAY()<(15-WEEKDAY(TODAY())))`;
    case 'lastMonth':
      return `AND(MONTH(${cell})=MONTH(EDATE(TODAY(),0-1)),YEAR(${cell})=YEAR(EDATE(TODAY(),0-1)))`;
    case 'thisMonth':
      return `AND(MONTH(${cell})=MONTH(TODAY()),YEAR(${cell})=YEAR(TODAY()))`;
    case 'nextMonth':
      return `AND(MONTH(${cell})=MONTH(EDATE(TODAY(),0+1)),YEAR(${cell})=YEAR(EDATE(TODAY(),0+1)))`;
    default:
      return undefined;
  }
}

/**
 * The cell the formula of a "dates occurring" rule is written for, as the formula spells it
 * ("B2", "$B2", ...): the first reference of the formula, when that is the first cell of the
 * rule and nothing else (no range, no sheet name). Whatever else the formula says does not
 * matter: a formula of another producer, or one with spaces of its own, names its cell all
 * the same. The plain name of the cell for any other formula, and for a rule without one.
 */
function periodRuleCell(rule: Pick<ConditionalRule, 'formulas'>, anchor: { row: number; col: number }): string {
  const plain = cellText(anchor.row, anchor.col);
  const written = typeof rule.formulas[0] === 'string' ? rule.formulas[0] : undefined;
  if (written === undefined) return plain;
  const first = formulaReferences(written)[0];
  if (!first || first.kind !== 'cell' || first.start !== first.bodyStart || first.write(false, false) !== plain) return plain;
  return first.write(first.absCol, first.absRow);
}

/** The thresholds a new icon set starts with, as Excel's do: the range shared out evenly, in percent. */
function iconThresholds(count: number): ConditionalValue[] {
  const out: ConditionalValue[] = [];
  for (let i = 0; i < count; i++) out.push({ type: 'percent', value: String(count === 3 ? [0, 33, 67][i] : Math.round((100 * i) / count)), gte: true });
  return out;
}

/** Kinds of rule that show what they find by themselves (a colour, a bar, an icon) and give no format. */
const formatsItself = (type: string) => type === 'colorScale' || type === 'dataBar' || type === 'iconSet';

/**
 * The rule the settings stand for, for a rule whose first cell is `anchor`. `base` is the rule
 * that is changed. While the kind of rule stays, the settings are a patch on it: what they do
 * not state, and what they state as the rule has it, stays as the rule has it - with everything
 * the settings have no word for (the thresholds of a scale, a bar or an icon set, the lengths of
 * a bar, a number of deviations, the spelling of a formula). Its place in the file and the record of
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
    case 'iconSet': {
      const kept = same && same.iconSet ? same.iconSet : undefined;
      const name = settings.iconSet !== undefined ? settings.iconSet : kept ? kept.set : '3TrafficLights1';
      const reverse = settings.reverse !== undefined ? settings.reverse === true : kept ? kept.reverse === true : false;
      const showValue = settings.showValue !== undefined ? settings.showValue !== false : kept ? kept.showValue !== false : true;
      if (kept && kept.set === name) {
        // The set it has, whatever that is: its thresholds stay as the file states them.
        if (kept.reverse !== reverse || kept.showValue !== showValue) rule.iconSet = { ...kept, reverse, showValue };
        break;
      }
      const icons = (ICON_SET_NAMES as readonly string[]).indexOf(name) >= 0 ? ICON_SETS.get(name) : undefined;
      if (!icons) throw new RangeError('This icon set is not known, or is not made here');
      // As many icons as there were: where one ends and the next begins stays. Icons chosen one by one are of the set that goes.
      const values = kept && !kept.custom && kept.values.length === icons.length ? kept.values : iconThresholds(icons.length);
      rule.iconSet = { set: name, values, showValue, reverse };
      break;
    }
    case 'timePeriod': {
      const period = settings.timePeriod !== undefined ? settings.timePeriod : same && typeof same.timePeriod === 'string' ? same.timePeriod : undefined;
      if (period === undefined) throw new RangeError('The rule needs a period');
      // The period the rule has: the rule as it is, with the formula as it is written.
      if (same && same.timePeriod === period) break;
      // Written for the cell the old formula named, the way it named it ("$A1").
      const formula = (TIME_PERIODS as readonly string[]).indexOf(period) >= 0 ? periodFormula(period, same ? periodRuleCell(same, anchor) : cell) : undefined;
      if (formula === undefined) throw new RangeError('This period is not known');
      rule.timePeriod = period;
      rule.formulas = [formula];
      break;
    }
    default:
    // duplicateValues and uniqueValues say everything by their kind.
  }
  // Scales, bars and icons show what they find by themselves: a rule that becomes one, and a new one, has no format. One that is
  // one already keeps the record a file names on it (nobody shows it, and nobody is to take it away).
  if (formatsItself(type)) return same ? rule : withFormat(rule, null, styles, read);
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
  if (patch.format !== undefined && !(patch.rule !== undefined && patch.rule.format !== undefined) && !formatsItself(rule.type)) rule = withFormat(rule, patch.format, workbook.styles, read);
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

// ---------------------------------------------------------------------------
// Rules that go with copied cells

/** A data validation of copied cells, cut down to those cells. */
export interface CopiedValidation {
  /** The cells of the copy the rule applies to: rows and columns count from the first row of the copy and its first cell. */
  ranges: CellRange[];
  /** The cell of the sheet the formulas are written for: the one the first of `ranges` starts at. */
  anchor: { row: number; col: number };
  rule: Omit<DataValidation, 'ranges'>;
  /** What the rule says, whatever cells it has (see pasteRules()): two rules with the same key are the same rule. */
  key: string;
}

/** The rules of one conditional format of copied cells, cut down to those cells. */
export interface CopiedConditional {
  /** As CopiedValidation.ranges. */
  ranges: CellRange[];
  anchor: { row: number; col: number };
  /** The rules as they applied when they were copied, with what ties them to the file: a paste that adds them takes that off. */
  rules: ConditionalRule[];
  /** As CopiedValidation.key, one for each of `rules`. */
  keys: string[];
}

/**
 * A rule of copied cells that a copy cannot take along: one that only the extension list of
 * the file has, or one with a formula that cannot be read to its end. Such a rule is kept as
 * it is wherever it is; a paste knows by this what the cells it writes ought to have had.
 */
export interface CopiedLeft {
  kind: 'validation' | 'conditional';
  ranges: CellRange[];
  key: string;
}

/**
 * The data validations and conditional formats of copied cells, as they applied when the
 * cells were copied. `left` are the rules of the cells that are not among them.
 */
export interface CopiedRules {
  validations: CopiedValidation[];
  conditional: CopiedConditional[];
  left: CopiedLeft[];
}

/** A formula of a rule that a live rule can hold: one whose sheet names can be read (see checkLiveFormulas()). */
const followsSheets = (text: string | undefined) => typeof text !== 'string' || text.indexOf('!') < 0 || !formulaNamesSheetsUnread(text);

/** A data validation that is kept as it is: one of the extension list, or one that could not follow the sheets it names. */
const fixedValidation = (now: DataValidation) => now.extended === true || !followsSheets(now.formula1) || !followsSheets(now.formula2);
/** The same for a conditional format: one rule of that kind keeps the whole format as it is. */
const fixedFormat = (now: ConditionalFormat) => now.rules.some((rule) => !!rule && (rule.extended === true || !ruleFormulaTexts(rule).every(followsSheets)));

const VALIDATION_FORMULAS = ['formula1', 'formula2'] as const;

/**
 * What a rule says, apart from the cells it has: its formulas as they would be written for
 * A1. Two rules of a sheet with the same key say the same for every cell.
 */
function validationKey(rule: Omit<DataValidation, 'ranges'>, anchor: { row: number; col: number }): string {
  return stable(SETTING_KEYS.map((key) => (key === 'formula1' || key === 'formula2' ? (typeof rule[key] === 'string' ? shiftFormula(rule[key] as string, -anchor.row, -anchor.col, true) : null) : rule[key] === undefined ? null : rule[key])).concat([rule.extended === true]));
}

/** The same for a conditional rule: all of it but its place in the file and in the order. */
function conditionalKey(rule: ConditionalRule, anchor: { row: number; col: number }): string {
  return stable({ ...rule, src: undefined, priority: 0, formulas: rule.formulas.map((text) => shiftFormula(text, -anchor.row, -anchor.col, true)) });
}

const firstCell = (ranges: readonly CellRange[]) => ({ row: ranges[0].top, col: ranges[0].left });

/**
 * Ranges that cover the same cells with fewer rectangles: two that share a whole side become
 * one, until none do. Each stands where the first of its parts stood in the list. The ranges
 * are taken not to overlap.
 */
export function joinedRanges(ranges: readonly CellRange[]): CellRange[] {
  if (ranges.length < 2) return ranges.slice();
  let list = ranges.map((range, at) => ({ top: range.top, left: range.left, bottom: range.bottom, right: range.right, at }));
  for (let joined = true; joined; ) {
    joined = false;
    for (const down of [true, false]) {
      list.sort(down ? (a, b) => a.left - b.left || a.right - b.right || a.top - b.top : (a, b) => a.top - b.top || a.bottom - b.bottom || a.left - b.left);
      const out: typeof list = [];
      for (const item of list) {
        const last = out[out.length - 1];
        if (last && (down ? last.left === item.left && last.right === item.right && last.bottom + 1 === item.top : last.top === item.top && last.bottom === item.bottom && last.right + 1 === item.left)) {
          if (down) last.bottom = item.bottom;
          else last.right = item.right;
          if (item.at < last.at) last.at = item.at;
          joined = true;
        } else out.push(item);
      }
      list = out;
    }
  }
  return list.sort((a, b) => a.at - b.at).map((item) => ({ top: item.top, left: item.left, bottom: item.bottom, right: item.right }));
}

/**
 * The part of ranges of the sheet that is in a copy, in the rows and columns of the copy, and
 * the cell of the sheet its first piece starts at. `runs` are the rows of the sheet that were
 * copied, in order, each with the row of the copy it starts at (rows that a filter hides are
 * left out of a copy). Undefined when no cell of the ranges was copied.
 */
function clippedToCopy(ranges: readonly CellRange[], left: number, right: number, runs: readonly { top: number; bottom: number; at: number }[]): { ranges: CellRange[]; anchor: { row: number; col: number } } | undefined {
  const out: CellRange[] = [];
  let anchor: { row: number; col: number } | undefined;
  for (const range of ranges) {
    const from = Math.max(range.left, left);
    const to = Math.min(range.right, right);
    if (from > to) continue;
    let last: CellRange | undefined;
    for (const run of runs) {
      if (run.top > range.bottom) break;
      const top = Math.max(range.top, run.top);
      const bottom = Math.min(range.bottom, run.bottom);
      if (top > bottom) continue;
      if (!anchor) anchor = { row: top, col: from };
      const piece = { top: run.at + (top - run.top), left: from - left, bottom: run.at + (bottom - run.top), right: to - left };
      // Rows on either side of rows that were left out stand next to each other in the copy: a range gives one piece.
      if (last && last.bottom + 1 === piece.top) last.bottom = piece.bottom;
      else out.push((last = piece));
    }
  }
  return anchor ? { ranges: out, anchor } : undefined;
}

/**
 * The rules that go with a copy of `range` (a selection of the sheet): the data validations
 * and the conditional formats that apply to its cells now, each cut down to those cells and
 * with its formulas written for the first of them. `hidden` are the rows of the range that
 * are not copied (sorted). Undefined when the cells have no rule.
 */
export function copyRules(workbook: Workbook, sheetIndex: number, range: CellRange, hidden: readonly number[] = []): CopiedRules | undefined {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet || ((!sheet.dataValidations || sheet.dataValidations.length === 0) && (!sheet.conditionalFormats || sheet.conditionalFormats.length === 0))) return undefined;
  const top = Math.max(0, range.top);
  const left = Math.max(0, range.left);
  const bottom = Math.min(range.bottom, MAX_ROWS - 1);
  const right = Math.min(range.right, MAX_COLS - 1);
  if (top > bottom || left > right) return undefined;
  const runs: { top: number; bottom: number; at: number }[] = [];
  let at = 0;
  let next = top;
  const run = (last: number) => {
    if (last < next) return;
    runs.push({ top: next, bottom: last, at });
    at += last - next + 1;
  };
  for (const row of hidden) {
    if (row < next || row > bottom) continue;
    run(row - 1);
    next = row + 1;
  }
  run(bottom);

  const space = ruleSpace(workbook, sheet);
  const selection = [{ top, left, bottom, right }];
  const copied: CopiedRules = { validations: [], conditional: [], left: [] };
  for (const rule of sheet.dataValidations || []) {
    const now = dataValidationNow(rule, space);
    if (!now || !rangesIntersect(now.ranges, selection)) continue;
    const part = clippedToCopy(now.ranges, left, right, runs);
    if (!part) continue;
    const key = validationKey(now, firstCell(now.ranges));
    if (fixedValidation(now)) {
      copied.left.push({ kind: 'validation', ranges: part.ranges, key });
      continue;
    }
    const dRow = part.anchor.row - now.ranges[0].top;
    const dCol = part.anchor.col - now.ranges[0].left;
    const taken = {} as Record<string, unknown>;
    for (const name of SETTING_KEYS) if (now[name] !== undefined) taken[name] = name === 'formula1' || name === 'formula2' ? shiftFormula(now[name] as string, dRow, dCol, true) : now[name];
    copied.validations.push({ ranges: part.ranges, anchor: part.anchor, rule: taken as unknown as Omit<DataValidation, 'ranges'>, key });
  }
  for (const format of sheet.conditionalFormats || []) {
    const now = conditionalFormatNow(format, space);
    if (!now || !rangesIntersect(now.ranges, selection)) continue;
    const part = clippedToCopy(now.ranges, left, right, runs);
    if (!part) continue;
    const rules = now.rules.filter((rule) => !!rule);
    const keys = rules.map((rule) => conditionalKey(rule, firstCell(now.ranges)));
    if (fixedFormat(now)) {
      for (const key of keys) copied.left.push({ kind: 'conditional', ranges: part.ranges, key });
      continue;
    }
    const to = { top: part.anchor.row, left: part.anchor.col, bottom: part.anchor.row, right: part.anchor.col };
    if (rules.length > 0) copied.conditional.push({ ranges: part.ranges, anchor: part.anchor, rules: rules.map((rule) => reanchored(rule, now.ranges[0], to)), keys });
  }
  return copied.validations.length > 0 || copied.conditional.length > 0 || copied.left.length > 0 ? copied : undefined;
}

/** How many rows and columns the rules of a copy reach over, counted from its first cell. */
export function copiedRulesExtent(copied: CopiedRules | undefined): { height: number; width: number } {
  let height = 0;
  let width = 0;
  if (copied) {
    for (const item of (copied.validations as { ranges: CellRange[] }[]).concat(copied.conditional, copied.left)) {
      for (const range of item.ranges) {
        if (range.bottom >= height) height = range.bottom + 1;
        if (range.right >= width) width = range.right + 1;
      }
    }
  }
  return { height, width };
}

const TOO_MANY_RANGES = 'The rule would be left with too many separate ranges';
const tooManyRanges = (error: unknown) => error instanceof RangeError && error.message === TOO_MANY_RANGES;

/**
 * Where ranges of a copy are once the copy is pasted over `area`, a block of `height` rows
 * and `width` columns repeated from the first cell of the area on: one range where the
 * repeats of a range touch. Cells beyond the area are left out. Undefined for more ranges
 * than a rule holds.
 */
function tiledRanges(ranges: readonly CellRange[], area: CellRange, height: number, width: number): CellRange[] | undefined {
  // A copy that is pasted once has the ranges it has, however many.
  const limit = Math.max(MAX_RULE_RANGES, ranges.length);
  const spans = (from: number, to: number, start: number, end: number, step: number): [number, number][] | undefined => {
    const out: [number, number][] = [];
    for (let at = start; at + from <= end; at += Math.max(1, step)) {
      const first = at + from;
      const last = Math.min(at + to, end);
      const before = out[out.length - 1];
      if (before && before[1] + 1 >= first) before[1] = Math.max(before[1], last);
      else out.push([first, last]);
      if (out.length > limit) return undefined;
    }
    return out;
  };
  let out: CellRange[] = [];
  for (const range of ranges) {
    const rows = spans(range.top, range.bottom, area.top, area.bottom, height);
    const cols = spans(range.left, range.right, area.left, area.right, width);
    if (!rows || !cols || rows.length * cols.length + out.length > limit) return undefined;
    const pieces: CellRange[] = [];
    for (const [top, bottom] of rows) for (const [left, right] of cols) pieces.push({ top, left, bottom, right });
    // Cells a rule names twice are in it once.
    if (out.length === 0) out = pieces;
    else {
      try {
        out = out.concat(subtractRanges(pieces, out));
      } catch (error) {
        if (!tooManyRanges(error)) throw error;
        return undefined;
      }
    }
  }
  return out;
}

/** Whether every cell of `ranges` is a cell of `within`. */
const rangesWithin = (ranges: readonly CellRange[], within: readonly CellRange[]) => subtractRanges(ranges, within).length === 0;

/** The cells of `ranges` that are cells of `areas`. */
function rangesInside(ranges: readonly CellRange[], areas: readonly CellRange[]): CellRange[] {
  const out: CellRange[] = [];
  for (const range of ranges) {
    for (const area of areas) {
      if (!intersects(range, area)) continue;
      out.push({ top: Math.max(range.top, area.top), left: Math.max(range.left, area.left), bottom: Math.min(range.bottom, area.bottom), right: Math.min(range.right, area.right) });
    }
  }
  return out;
}

/** A rule that stays where it is, whatever is pasted: what it says, the cells it has, and those of them that are pasted over. */
interface Staying {
  key: string;
  ranges: CellRange[];
  inside: CellRange[];
}

/**
 * How many rules a paste could not do as they are where the cells come from: rules of the
 * copy that did not come (`left`, placed where the copy is pasted) and rules that stayed
 * though the cells they are on are copies of cells without them. A rule that stayed on cells
 * whose copy has that very rule is neither.
 */
function countUndone(staying: readonly Staying[], left: readonly { key: string; ranges: CellRange[] | undefined }[]): number {
  let count = 0;
  const within = (ranges: readonly CellRange[], others: readonly CellRange[]) => {
    try {
      return rangesWithin(ranges, others);
    } catch (error) {
      if (!tooManyRanges(error)) throw error;
      return false;
    }
  };
  for (const rule of staying) if (!left.some((item) => item.key === rule.key && !!item.ranges && within(rule.inside, item.ranges))) count++;
  for (const item of left) if (!item.ranges || !staying.some((rule) => rule.key === item.key && within(item.ranges as CellRange[], rule.ranges))) count++;
  return count;
}

/** The rules a copy left behind, of one kind, where the copy is pasted; without ranges for a rule that would have too many. */
function leftPlaced(copied: CopiedRules | undefined, kind: CopiedLeft['kind'], areas: readonly CellRange[], height: number, width: number): { key: string; ranges: CellRange[] | undefined }[] {
  const out: { key: string; ranges: CellRange[] | undefined }[] = [];
  for (const item of copied ? copied.left : []) if (item.kind === kind) out.push({ key: item.key, ranges: tiledRanges(item.ranges, areas[0], height, width) });
  return out;
}

/** The rule of the file a changed rule came from, when the rule says again what that one says for the same cells; undefined otherwise. */
function validationAsRead(sheet: Sheet, rule: DataValidation, space: RuleSpace): DataValidation | undefined {
  if (rule.src === undefined || !sheet.validationsRead) return undefined;
  for (const read of sheet.validationsRead.list || []) {
    if (!read || read.live || read.src !== rule.src) continue;
    const now = dataValidationNow(read, space);
    return now && sameRanges(now.ranges, rule.ranges) && sameSettings(now, rule) && now.extended === rule.extended ? read : undefined;
  }
  return undefined;
}

/** The same for a conditional format all of whose rules came from one element of the file. */
function formatAsRead(sheet: Sheet, format: ConditionalFormat, space: RuleSpace): ConditionalFormat | undefined {
  const first = format.rules[0];
  if (!first || !first.src || !sheet.conditionalRead) return undefined;
  for (const read of sheet.conditionalRead.list || []) {
    if (!read || read.live || read.rules.length !== format.rules.length || !read.rules[0] || !read.rules[0].src || read.rules[0].src[0] !== first.src[0]) continue;
    const now = conditionalFormatNow(read, space);
    if (!now || !sameRanges(now.ranges, format.ranges)) return undefined;
    // Its place in the order is given anew when the list is put together.
    return now.rules.every((rule, i) => !!rule && !!format.rules[i] && stable({ ...rule, priority: 0 }) === stable({ ...format.rules[i], priority: 0 })) ? read : undefined;
  }
  return undefined;
}

/**
 * The data validations of a sheet once a copy with the rules `wanted` is pasted over `areas`
 * (without a copy: once the cells of the areas are taken out of every rule), and how many
 * rules could not be done (see countUndone()). The list of the sheet itself when no rule changes.
 */
function validationsPasted(workbook: Workbook, sheet: Sheet, copied: CopiedRules | undefined, areas: readonly CellRange[], height: number, width: number): { list: DataValidation[] | undefined; skipped: number } {
  const space = ruleSpace(workbook, sheet);
  const current = sheet.dataValidations || [];
  const wanted = copied ? copied.validations : [];
  const left = leftPlaced(copied, 'validation', areas, height, width);
  if (current.length === 0 && wanted.length === 0) return { list: sheet.dataValidations, skipped: left.length };
  const nows = current.map((rule) => dataValidationNow(rule, space));
  // The cells of a rule that is kept as it is keep that rule, and get no other.
  const staying: Staying[] = [];
  let fixed: CellRange[] = [];
  for (const now of nows) {
    if (!now || !fixedValidation(now) || !rangesIntersect(now.ranges, areas)) continue;
    staying.push({ key: validationKey(now, firstCell(now.ranges)), ranges: now.ranges, inside: rangesInside(now.ranges, areas) });
    fixed = fixed.concat(now.ranges);
  }
  let skipped = countUndone(staying, left);
  const free = fixed.length > 0 ? subtractRanges(areas, fixed) : areas.slice();
  // A cell has one rule: the first of the copy that names it.
  let claimed: CellRange[] = [];
  const pasted: { ranges: CellRange[]; rule: Omit<DataValidation, 'ranges'>; key: string; taken: boolean }[] = [];
  for (const item of wanted) {
    const all = tiledRanges(item.ranges, areas[0], height, width);
    if (!all) {
      skipped++;
      continue;
    }
    const ranges = fixed.length > 0 || claimed.length > 0 ? subtractRanges(all, fixed.concat(claimed)) : all;
    if (ranges.length === 0) continue;
    claimed = claimed.concat(ranges);
    const rule = { ...item.rule };
    for (const name of VALIDATION_FORMULAS) if (typeof rule[name] === 'string') rule[name] = shiftFormula(rule[name] as string, ranges[0].top - item.anchor.row, ranges[0].left - item.anchor.col, true);
    pasted.push({ ranges, rule, key: item.key, taken: false });
  }
  if (free.length === 0) return { list: sheet.dataValidations, skipped };

  // The rule the copy brings may be one the sheet has (for the cells that were copied, say): that one then has the
  // cells as well, as in Excel, and no second rule says the same. A rule that has exactly these cells comes first.
  const byKey = new Map<string, (typeof pasted)[number][]>();
  for (const item of pasted) {
    const list = byKey.get(item.key);
    if (list) list.push(item);
    else byKey.set(item.key, [item]);
  }
  const parts = nows.map((now) => {
    if (!now || fixedValidation(now)) return undefined;
    const touched = rangesIntersect(now.ranges, free);
    const candidates = byKey.size > 0 ? byKey.get(validationKey(now, firstCell(now.ranges))) : undefined;
    if (!touched && !candidates) return undefined;
    const kept = touched ? subtractRanges(now.ranges, free) : now.ranges;
    return { now, touched, kept, inside: touched ? rangesInside(now.ranges, free) : [], candidates, same: undefined as (typeof pasted)[number] | undefined, exact: false };
  });
  for (const exact of [true, false]) {
    for (const part of parts) {
      if (!part || part.same || !part.candidates || (exact && !part.touched)) continue;
      for (const item of part.candidates) {
        if (item.taken || (exact && !(rangesWithin(part.inside, item.ranges) && rangesWithin(item.ranges, part.inside)))) continue;
        item.taken = true;
        part.exact = exact;
        part.same = item;
        break;
      }
    }
  }
  const changed = new Map<number, DataValidation | undefined>();
  parts.forEach((part, at) => {
    // The cells it has there are the cells it gets: it stays the rule it is, as the file has it.
    if (!part || part.exact || (!part.touched && !part.same)) return;
    const now = part.now;
    const ranges = joinedRanges(part.same ? part.kept.concat(part.same.ranges) : part.kept);
    if (ranges.length === 0) {
      changed.set(at, undefined);
      return;
    }
    if (ranges.length > Math.max(MAX_RULE_RANGES, now.ranges.length)) throw new RangeError(TOO_MANY_RANGES);
    const next: DataValidation = { ...now, ranges, live: true };
    // The formulas are written for the first cell of the rule, which may be another one now.
    const dRow = ranges[0].top - now.ranges[0].top;
    const dCol = ranges[0].left - now.ranges[0].left;
    if (dRow !== 0 || dCol !== 0) for (const name of VALIDATION_FORMULAS) if (typeof next[name] === 'string') next[name] = shiftFormula(next[name] as string, dRow, dCol, true);
    // Cut and pasted back, it is the rule of the file again.
    changed.set(at, validationAsRead(sheet, next, space) || next);
  });
  const added = pasted.filter((item) => !item.taken);
  if (changed.size === 0 && added.length === 0) return { list: sheet.dataValidations, skipped };
  const next: DataValidation[] = [];
  current.forEach((rule, at) => {
    if (!changed.has(at)) next.push(rule);
    else if (changed.get(at)) next.push(changed.get(at) as DataValidation);
  });
  // Behind the rules of the main part and before those of the extension list, which is where a file has them.
  let at = next.length;
  while (at > 0 && next[at - 1].extended) at--;
  next.splice(at, 0, ...added.map((item): DataValidation => ({ ...item.rule, ranges: item.ranges, live: true })));
  if (next.length > MAX_SHEET_VALIDATIONS) throw new RangeError('The sheet has too many data validations');
  return { list: next.length > 0 ? next : undefined, skipped };
}

/**
 * The operation that makes the conditional formats of a sheet what they are once a copy with
 * the formats of `copied` is pasted over `areas` (without a copy: once the cells of the areas
 * are taken out of every rule), and how many rules could not be done (see countUndone()). A
 * format that is kept as it is (see fixedFormat()) stays with all its cells.
 *
 * What holds afterwards: the rules that apply to a cell that was pasted over are, in the order
 * they are applied in, those of the cell it is a copy of; and for no other cell does the order
 * of its rules change. A rule the sheet has is given the cells only where that holds.
 */
function conditionalPasted(workbook: Workbook, sheetIndex: number, sheet: Sheet, copied: CopiedRules | undefined, areas: readonly CellRange[], height: number, width: number): { ops: Op[]; skipped: number } {
  const current = sheet.conditionalFormats || [];
  const wanted = copied ? copied.conditional : [];
  const left = leftPlaced(copied, 'conditional', areas, height, width);
  if (current.length === 0 && wanted.length === 0) return { ops: [], skipped: left.length };
  let skipped = 0;
  // The rules of the copy one by one, in the order they are applied in where they come from; those of one format share their ranges.
  let pasted: { ranges: CellRange[]; rule: ConditionalRule; key: string; same: string | undefined }[] = [];
  for (const item of wanted) {
    const ranges = tiledRanges(item.ranges, areas[0], height, width);
    if (!ranges) {
      skipped += item.rules.length;
      continue;
    }
    if (ranges.length === 0) continue;
    const from = { top: item.anchor.row, left: item.anchor.col, bottom: item.anchor.row, right: item.anchor.col };
    item.rules.forEach((rule, i) => pasted.push({ ranges, rule: reanchored(rule, from, ranges[0]), key: item.keys[i], same: undefined }));
  }
  pasted = pasted
    .map((item, order) => ({ item, order }))
    .sort((a, b) => {
      const pa = priorityOf(a.item.rule);
      const pb = priorityOf(b.item.rule);
      return pa === pb ? a.order - b.order : pa < pb ? -1 : 1;
    })
    .map((entry) => entry.item);

  const slots = slotsInOrder(workbook, sheet);
  const renumber = canRenumber(slots);
  const space = ruleSpace(workbook, sheet);
  const staying: Staying[] = [];
  const parts = current.map((format) => {
    const now = conditionalFormatNow(format, space);
    if (!now) return undefined;
    const touched = rangesIntersect(now.ranges, areas);
    if (fixedFormat(now)) {
      if (touched) for (const rule of now.rules) if (rule) staying.push({ key: conditionalKey(rule, firstCell(now.ranges)), ranges: now.ranges, inside: rangesInside(now.ranges, areas) });
      return undefined;
    }
    return { now, touched, kept: undefined as CellRange[] | undefined, inside: touched ? rangesInside(now.ranges, areas) : [] };
  });
  skipped += countUndone(staying, left);

  // The rule the copy brings may be one the sheet has (for the cells that were copied, say): that one then has the
  // cells as well, as in Excel, and no second rule says the same. A rule that has exactly these cells comes first.
  const order = new Map<string, number>();
  const taken = new Map<string, (typeof pasted)[number]>();
  if (pasted.length > 0) {
    const byKey = new Map<string, { place: string; f: number }[]>();
    slots.forEach((slot, at) => {
      const place = `${slot.f}:${slot.r}`;
      order.set(place, at);
      if (!parts[slot.f]) return;
      const key = conditionalKey(slot.rule, firstCell(slot.ranges));
      const list = byKey.get(key);
      if (list) list.push({ place, f: slot.f });
      else byKey.set(key, [{ place, f: slot.f }]);
    });
    for (const exact of [true, false]) {
      for (const item of pasted) {
        if (item.same !== undefined) continue;
        for (const found of byKey.get(item.key) || []) {
          if (taken.has(found.place)) continue;
          if (exact) {
            const part = parts[found.f] as { touched: boolean; inside: CellRange[] };
            if (!part.touched || !(rangesWithin(part.inside, item.ranges) && rangesWithin(item.ranges, part.inside))) continue;
          }
          item.same = found.place;
          taken.set(found.place, item);
          break;
        }
      }
    }
    // The rules that are new to the sheet are applied before all it has (after them in a sheet whose numbers stay).
    // A rule of the sheet gets the cells only if every rule of the copy that shares a cell with it then still comes
    // before or after it as it does in the copy.
    const at = (item: (typeof pasted)[number]) => order.get(item.same as string) as number;
    const drop = (item: (typeof pasted)[number]) => {
      taken.delete(item.same as string);
      item.same = undefined;
    };
    if (taken.size > 0) {
      if (renumber) {
        for (let j = pasted.length - 1; j >= 0; j--) {
          const item = pasted[j];
          if (item.same === undefined) continue;
          for (let l = j + 1; l < pasted.length; l++) {
            const later = pasted[l];
            if ((later.same !== undefined && at(item) < at(later)) || !(later.ranges === item.ranges || rangesIntersect(later.ranges, item.ranges))) continue;
            drop(item);
            break;
          }
        }
      } else {
        for (let j = 0; j < pasted.length; j++) {
          const item = pasted[j];
          if (item.same === undefined) continue;
          for (let i = 0; i < j; i++) {
            const earlier = pasted[i];
            if ((earlier.same !== undefined && at(earlier) < at(item)) || !(earlier.ranges === item.ranges || rangesIntersect(earlier.ranges, item.ranges))) continue;
            drop(item);
            break;
          }
        }
      }
    }
  }

  const replaced = new Map<number, ConditionalFormat>();
  const gone = new Set<string>();
  const moved = new Map<string, Placed>();
  const sized = (ranges: CellRange[], had: number) => {
    if (ranges.length > Math.max(MAX_RULE_RANGES, had)) throw new RangeError(TOO_MANY_RANGES);
    return ranges;
  };
  parts.forEach((part, f) => {
    if (!part) return;
    const now = part.now;
    const rules = now.rules.map((rule, r) => ({ rule, r, same: taken.get(`${f}:${r}`) })).filter((entry) => !!entry.rule);
    const found = rules.filter((entry) => entry.same);
    if (rules.length === 0 || (found.length === 0 && !part.touched)) return;
    const kept = part.touched ? subtractRanges(now.ranges, areas) : now.ranges;
    const whole = (ranges: CellRange[]): ConditionalFormat => {
      const format: ConditionalFormat = { ...now, ranges, rules: now.rules.map((rule) => (rule ? reanchored(rule, now.ranges[0], ranges[0]) : rule)), live: true };
      // Cut and pasted back, it is the format of the file again.
      return formatAsRead(sheet, format, space) || format;
    };
    const first = found.length > 0 ? (found[0].same as (typeof pasted)[number]).ranges : [];
    if (found.length === rules.length && found.every((entry) => (entry.same as (typeof pasted)[number]).ranges === first)) {
      // All its rules get the same cells: the format stays one format. With the cells it has there already, it stays as the file has it.
      if (part.touched && rangesWithin(part.inside, first) && rangesWithin(first, part.inside)) return;
      replaced.set(f, whole(sized(joinedRanges(kept.concat(first)), now.ranges.length)));
      return;
    }
    // Rule by rule: one that gets cells of the copy becomes a format of its own, the others keep what is left.
    for (const entry of rules) {
      const key = `${f}:${entry.r}`;
      if (entry.same) {
        const ranges = sized(joinedRanges(kept.concat(entry.same.ranges)), now.ranges.length);
        moved.set(key, { f, r: entry.r, rule: reanchored(entry.rule, now.ranges[0], ranges[0]), detach: ranges });
      } else if (kept.length === 0) gone.add(key);
    }
    if (part.touched && kept.length > 0 && found.length < rules.length) replaced.set(f, whole(kept));
  });
  // The rules that are new to the sheet, in the order of the copy.
  const added: Placed[] = [];
  for (const item of pasted) {
    if (item.same !== undefined) continue;
    const plain = mainPartRule(item.rule);
    if (!plain) continue;
    // What the extension list says about the rule it was copied from stays with that one.
    if (item.rule.extId !== undefined || (item.rule.iconSet && item.rule.iconSet.custom)) skipped++;
    added.push({ f: -1, r: -1, rule: plain, detach: item.ranges });
  }
  if (replaced.size === 0 && gone.size === 0 && moved.size === 0 && added.length === 0) return { ops: [], skipped };
  const removed = slots.filter((slot) => gone.has(`${slot.f}:${slot.r}`));
  const stayed = slots
    .filter((slot) => !gone.has(`${slot.f}:${slot.r}`))
    .map((slot): Placed => {
      const format = replaced.get(slot.f);
      return moved.get(`${slot.f}:${slot.r}`) || (format ? { f: slot.f, r: slot.r, rule: format.rules[slot.r] } : stay(slot, sheet));
    });
  // As a rule that is added, the rules of the copy are the first to be applied - except in a sheet whose numbers stay.
  // (UNSURE: where Excel puts pasted rules in the order was not looked up.)
  if (renumber) return { ops: conditionalOp(sheetIndex, workbook, sheet, added.concat(stayed), true, removed, replaced), skipped };
  let highest = 0;
  for (const format of current) for (const other of format.rules) if (other && other.priority < Infinity && other.priority > highest) highest = Math.floor(other.priority);
  for (const item of added) item.rule = { ...item.rule, priority: ++highest };
  return { ops: conditionalOp(sheetIndex, workbook, sheet, stayed.concat(added), false, removed, replaced), skipped };
}

/**
 * What a paste of copied cells does to the rules of the cells it writes: `area` is the range
 * that was pasted over, the copy a block of `height` rows and `width` columns that is repeated
 * from its first cell on. As Excel's "paste all" does, the cells get the data validation and
 * the conditional formats of the cells they are copies of - and lose their own, also where
 * the copy has none. (UNSURE: taken from how Excel is described to behave, not compared with
 * it. The same is not done for notes, see pasteNotes().) Formulas move as those of cells do:
 * relative references by the distance of the paste, anchored ones not at all. `area` is what
 * the cells, notes and rules of the copy reach over: where a selection was copied that goes on
 * beyond that (empty cells without a rule), the cells under the rest of it keep their rules
 * (UNSURE: Excel takes the whole selection).
 *
 * A rule the sheet has already (the same rule for the cells beside, say) is given the cells
 * instead of being made a second time, and a paste that changes no rule gives no operation.
 * Rules that are kept as they are - those only the extension list has, those with a formula
 * that cannot be read to its end - are neither made nor taken away: the cells under such a
 * data validation keep it, a conditional format with such a rule stays as it is. A rule that
 * would have more ranges than a rule holds is not pasted, and where the rules that are there
 * would be left with so many, the rules of that kind all stay as they are. None of this
 * refuses the paste: `skipped` counts the rules that could not be done as they are where the
 * cells come from - those of the copy that did not come (or came without what the extension
 * list says of them), and those that stayed though the copy has none there. Throws RangeError
 * only for a sheet that would hold more rules, or a workbook more formats, than a file can.
 */
export function pasteRules(workbook: Workbook, sheetIndex: number, copied: CopiedRules | undefined, area: CellRange, height: number, width: number): { ops: Op[]; skipped: number } {
  const sheet = workbook.sheets[sheetIndex];
  const all = copied ? copied.left.length + copied.validations.length + copied.conditional.reduce((sum, item) => sum + item.rules.length, 0) : 0;
  if (!sheet || sheet.kind !== 'worksheet') return { ops: [], skipped: all };
  const target = targetRanges(area);
  return rulesPasted(workbook, sheetIndex, sheet, copied, target, height, width);
}

function rulesPasted(workbook: Workbook, sheetIndex: number, sheet: Sheet, copied: CopiedRules | undefined, areas: CellRange[], height: number, width: number): { ops: Op[]; skipped: number } {
  let ops: Op[] = [];
  let skipped = 0;
  // Where the rules that are there cannot be cut (too many ranges would be left), they all stay: nothing is lost, and it is said.
  const undone = (kind: CopiedLeft['kind'], count: number) => Math.max(1, count + (copied ? copied.left.filter((item) => item.kind === kind).length : 0));
  try {
    const validations = validationsPasted(workbook, sheet, copied, areas, height, width);
    skipped += validations.skipped;
    if (validations.list !== sheet.dataValidations) ops.push({ type: 'setDataValidations', sheet: sheetIndex, validations: validations.list });
  } catch (error) {
    if (!tooManyRanges(error)) throw error;
    skipped += undone('validation', copied ? copied.validations.length : 0);
  }
  try {
    const conditional = conditionalPasted(workbook, sheetIndex, sheet, copied, areas, height, width);
    skipped += conditional.skipped;
    ops = ops.concat(conditional.ops);
  } catch (error) {
    if (!tooManyRanges(error)) throw error;
    skipped += undone('conditional', copied ? copied.conditional.reduce((sum, item) => sum + item.rules.length, 0) : 0);
  }
  return { ops, skipped };
}

/**
 * Takes the cells of ranges out of the rules of a sheet, as a cut does: their data validation
 * and their conditional formats go with them (a paste of the copy puts them where the cells
 * go). Rules that are kept as they are stay (see pasteRules()): the cells are cut all the
 * same, and no rule is lost.
 */
export function cutRules(workbook: Workbook, sheetIndex: number, ranges: readonly CellRange[]): Op[] {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet || sheet.kind !== 'worksheet' || ranges.length === 0) return [];
  if ((!sheet.dataValidations || sheet.dataValidations.length === 0) && (!sheet.conditionalFormats || sheet.conditionalFormats.length === 0)) return [];
  let target: CellRange[];
  try {
    target = targetRanges(ranges);
  } catch (error) {
    if (!(error instanceof RangeError)) throw error;
    return [];
  }
  return rulesPasted(workbook, sheetIndex, sheet, undefined, target, 1, 1).ops;
}
