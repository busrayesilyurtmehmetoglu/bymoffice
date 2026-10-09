// The dialogs that set rules on cells: data validation. A dialog is modal
// within the element it is put in: it covers that element, keeps the focus
// inside itself, closes on Escape and applies on Enter. Whatever comes from the
// file (the messages of a rule, its formulas) only ever reaches the page as the
// value of a field or through textContent.
//
// What is typed into a dialog is typed the way the locale writes it: numbers
// with its decimal separator, dates and times as it writes them, formulas with
// its separators (and, in Turkish, its function names), the items of a list
// separated by its list separator. validationFromFields() turns that into what
// a file stores, and validationFields() shows a rule of a file that way.
//
// Nothing in this module touches the DOM at import time.

import { parseFormula } from '../calc/parse';
import { cellEditText, parseCellInput } from '../edit/input';
import { formulaFromLocal, formulaToLocal } from '../edit/formula-locale';
import type { FormatLocale } from '../format/locale';
import { MAX_COLS, MAX_ROWS, parseCellRef, parseRange, rangeName, type CellRange, type Color, type ConditionalRule, type DataValidation } from '../model/sheet/types';
import {
  CONDITIONAL_TYPES,
  MAX_CONDITIONAL_TEXT,
  MAX_VALIDATION_LIST,
  MAX_VALIDATION_MESSAGE,
  MAX_VALIDATION_TITLE,
  VALIDATION_OPERATORS,
  VALIDATION_TYPES,
  ruleFormatOf,
  textRuleCell,
  type ConditionalPatch,
  type ConditionalRuleInfo,
  type ConditionalSettings,
  type ConditionalType,
  type RuleFormat,
  type ValidationOperator,
  type ValidationSettings,
  type ValidationType,
} from '../ops/rules';
import type { RuleStrings } from './strings';

// ---------------------------------------------------------------------------
// Fields and what they stand for

/** The fields of the data validation dialog, as they are typed. */
export interface ValidationFields {
  type: string;
  operator: string;
  /** The value, the lowest value, the source of a list or the formula of a custom rule. */
  value1: string;
  /** The highest value of "between" and "not between". */
  value2: string;
  allowBlank: boolean;
  /** A list shows its arrow in the cell. */
  inCellDropdown: boolean;
  showInputMessage: boolean;
  promptTitle: string;
  prompt: string;
  showErrorMessage: boolean;
  errorStyle: string;
  errorTitle: string;
  error: string;
}

/** What a field of a dialog cannot be taken as; `field` names it, so that the dialog can point at it. */
export class FieldError extends RangeError {
  readonly field: string;

  constructor(message: string, field: string) {
    super(message);
    this.name = 'FieldError';
    this.field = field;
  }
}

export interface FieldOptions {
  date1904?: boolean;
  /** Whether the workbook has a defined name of this spelling (see formulaFromLocal()). */
  isName?(name: string): boolean;
}

const PLAIN_NUMBER = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;

/** The separator of the items of a list as the locale types it: ";" where the comma is the decimal separator. */
export const listSeparator = (locale: FormatLocale): string => (locale.decimal === ',' ? ';' : ',');

/** The items of a list written into a rule ("a,b,c" in quotes); undefined for any other formula. */
function constantItems(formula: string): string[] | undefined {
  const text = formula.trim();
  if (text.length < 2 || text.charAt(0) !== '"' || text.charAt(text.length - 1) !== '"') return undefined;
  const inner = text.slice(1, -1);
  if (/(^|[^"])"([^"]|$)/.test(inner.replace(/""/g, ''))) return undefined;
  return inner
    .replace(/""/g, '"')
    .split(',')
    .map((item) => item.trim())
    .filter((item) => item !== '');
}

/** A stored value of a rule the way it is typed. */
function shownValue(text: string | undefined, type: string, locale: FormatLocale, date1904: boolean): string {
  if (typeof text !== 'string' || text === '') return '';
  if (type === 'list') {
    const items = constantItems(text);
    if (items) return items.join(`${listSeparator(locale)} `);
  } else if (type !== 'custom' && PLAIN_NUMBER.test(text.trim())) {
    const value = Number(text);
    if (isFinite(value)) {
      const dated = type === 'date' || type === 'time';
      // Any date format will do: it only says that the number is to be read as a date.
      return cellEditText({ kind: 'number', value, style: 0 }, dated ? 'yyyy-mm-dd' : 'General', { locale, date1904 });
    }
  }
  return `=${formulaToLocal(text, locale)}`;
}

/** The fields of the dialog for a rule; for no rule, the fields a new rule starts with. */
export function validationFields(rule: DataValidation | undefined, locale: FormatLocale, date1904: boolean): ValidationFields {
  if (!rule) {
    return {
      type: 'none',
      operator: 'between',
      value1: '',
      value2: '',
      allowBlank: true,
      inCellDropdown: true,
      showInputMessage: true,
      promptTitle: '',
      prompt: '',
      showErrorMessage: true,
      errorStyle: 'stop',
      errorTitle: '',
      error: '',
    };
  }
  const type = (VALIDATION_TYPES as readonly string[]).indexOf(rule.type) >= 0 ? rule.type : 'none';
  return {
    type,
    operator: (VALIDATION_OPERATORS as readonly string[]).indexOf(rule.operator) >= 0 ? rule.operator : 'between',
    value1: shownValue(rule.formula1, type, locale, date1904),
    value2: shownValue(rule.formula2, type, locale, date1904),
    allowBlank: rule.allowBlank === true,
    inCellDropdown: rule.hideDropDown !== true,
    showInputMessage: rule.showInputMessage === true,
    promptTitle: typeof rule.promptTitle === 'string' ? rule.promptTitle : '',
    prompt: typeof rule.prompt === 'string' ? rule.prompt : '',
    showErrorMessage: rule.showErrorMessage === true,
    errorStyle: rule.errorStyle === 'warning' || rule.errorStyle === 'information' ? rule.errorStyle : 'stop',
    errorTitle: typeof rule.errorTitle === 'string' ? rule.errorTitle : '',
    error: typeof rule.error === 'string' ? rule.error : '',
  };
}

/** The label of a value field, by the kind of rule and whether it is one of two. */
export function valueLabel(type: string, operator: string, second: boolean, strings: RuleStrings): string {
  const two = operator === 'between' || operator === 'notBetween';
  if (type === 'list') return strings.source;
  if (type === 'custom') return strings.formula;
  if (type === 'date') return two ? (second ? strings.endDate : strings.startDate) : strings.date;
  if (type === 'time') return two ? (second ? strings.endTime : strings.startTime) : strings.time;
  return two ? (second ? strings.maximum : strings.minimum) : strings.value;
}

/** Whether a kind of rule has an operator, and whether the operator has a second value. */
export const hasOperator = (type: string): boolean => type !== 'none' && type !== 'list' && type !== 'custom';
export const hasSecondValue = (type: string, operator: string): boolean => hasOperator(type) && (operator === 'between' || operator === 'notBetween');

/**
 * The rule the fields of the dialog stand for, as setValidation() takes it. Throws FieldError
 * (in the language of `strings`) for a field that cannot be taken: a value that is missing or
 * is not what the kind of rule asks for, a formula that cannot be read, a text that is too long.
 */
export function validationFromFields(fields: ValidationFields, locale: FormatLocale, strings: RuleStrings, options: FieldOptions = {}): ValidationSettings {
  const type = ((VALIDATION_TYPES as readonly string[]).indexOf(fields.type) >= 0 ? fields.type : 'none') as ValidationType;
  const date1904 = options.date1904 === true;
  const limited = (text: string, limit: number, field: string, label: string): string => {
    if (text.length > limit) throw new FieldError(strings.textTooLong(label, limit), field);
    return text;
  };
  const settings: ValidationSettings = {
    type,
    allowBlank: fields.allowBlank,
    hideDropDown: type === 'list' && !fields.inCellDropdown,
    showInputMessage: fields.showInputMessage,
    promptTitle: limited(fields.promptTitle, MAX_VALIDATION_TITLE, 'promptTitle', `${strings.inputMessage}, ${strings.title}`),
    prompt: limited(fields.prompt, MAX_VALIDATION_MESSAGE, 'prompt', `${strings.inputMessage}, ${strings.message}`),
    showErrorMessage: fields.showErrorMessage,
    errorStyle: fields.errorStyle === 'warning' || fields.errorStyle === 'information' ? fields.errorStyle : 'stop',
    errorTitle: limited(fields.errorTitle, MAX_VALIDATION_TITLE, 'errorTitle', `${strings.errorAlert}, ${strings.title}`),
    error: limited(fields.error, MAX_VALIDATION_MESSAGE, 'error', `${strings.errorAlert}, ${strings.message}`),
  };
  if (type === 'none') return settings;

  const operator = ((VALIDATION_OPERATORS as readonly string[]).indexOf(fields.operator) >= 0 ? fields.operator : 'between') as ValidationOperator;
  const formula = (text: string, field: string, label: string): string => {
    let stored: string;
    try {
      stored = formulaFromLocal(text, locale, { isName: options.isName });
    } catch (error) {
      throw new FieldError(strings.badRuleFormula(label), field);
    }
    if (stored.trim() === '') throw new FieldError(strings.needValue(label), field);
    try {
      parseFormula(stored);
    } catch (error) {
      throw new FieldError(strings.badRuleFormula(label), field);
    }
    return stored;
  };
  /** One value field as a file stores it. */
  const value = (typed: string, field: string, second: boolean): string => {
    const label = valueLabel(type, operator, second, strings);
    const text = typed.trim();
    if (text === '') throw new FieldError(type === 'list' ? strings.needItems : strings.needValue(label), field);
    if (text.charAt(0) === '=') return formula(text.slice(1), field, label);
    if (type === 'custom') return formula(text, field, label);
    if (type === 'list') {
      const separator = listSeparator(locale);
      const items = text
        .split(separator)
        .map((item) => item.trim())
        .filter((item) => item !== '');
      if (items.length === 0) throw new FieldError(strings.needItems, field);
      // The file separates the items with commas whatever the locale: an item cannot hold one.
      if (separator !== ',' && items.some((item) => item.indexOf(',') >= 0)) throw new FieldError(strings.itemWithComma, field);
      const joined = items.join(',');
      if (joined.length > MAX_VALIDATION_LIST) throw new FieldError(strings.listTooLong(MAX_VALIDATION_LIST), field);
      return `"${joined.replace(/"/g, '""')}"`;
    }
    const parsed = parseCellInput(text, { locale, date1904 });
    if (parsed.kind !== 'number' || typeof parsed.value !== 'number' || parsed.formula !== undefined || !isFinite(parsed.value)) {
      throw new FieldError(type === 'date' ? strings.needDate(label) : type === 'time' ? strings.needTime(label) : type === 'whole' ? strings.needWholeNumber(label) : strings.needNumber(label), field);
    }
    if (type === 'whole' && Math.floor(parsed.value) !== parsed.value) throw new FieldError(strings.needWholeNumber(label), field);
    return String(parsed.value);
  };
  settings.formula1 = value(fields.value1, 'value1', false);
  if (hasOperator(type)) {
    settings.operator = operator;
    if (hasSecondValue(type, operator)) settings.formula2 = value(fields.value2, 'value2', true);
  }
  return settings;
}

/**
 * What the dialog changes of the rule it was opened with: undefined when no field is other
 * than it was (OK is then no change at all), else the rule the fields stand for. A value
 * field that was not touched keeps what the rule holds, as the file wrote it: a list keeps
 * its spelling, a formula its spaces. Throws FieldError like validationFromFields().
 */
export function validationChange(
  fields: ValidationFields,
  initial: { fields: ValidationFields; rule: DataValidation | undefined },
  locale: FormatLocale,
  strings: RuleStrings,
  options: FieldOptions = {},
): ValidationSettings | undefined {
  const now = fields as unknown as Record<string, unknown>;
  const before = initial.fields as unknown as Record<string, unknown>;
  if (Object.keys(before).every((key) => now[key] === before[key])) return undefined;
  const rule = initial.rule;
  const sameKind = !!rule && fields.type === initial.fields.type;
  const keep1 = sameKind && fields.value1 === initial.fields.value1 && typeof rule.formula1 === 'string';
  const keep2 = sameKind && fields.value2 === initial.fields.value2 && typeof rule.formula2 === 'string';
  // A field that is kept is not read: any formula stands in for it while the others are.
  const settings = validationFromFields({ ...fields, value1: keep1 ? '=1' : fields.value1, value2: keep2 ? '=1' : fields.value2 }, locale, strings, options);
  if (keep1 && settings.formula1 !== undefined) settings.formula1 = (rule as DataValidation).formula1;
  if (keep2 && settings.formula2 !== undefined) settings.formula2 = (rule as DataValidation).formula2;
  return settings;
}

// ---------------------------------------------------------------------------
// Conditional formats: fields and what they stand for

/** The fields of the dialog of one conditional rule, as they are typed. Which of them count depends on the kind of rule. */
export interface ConditionalFields {
  type: string;
  /** cellIs. */
  operator: string;
  /** cellIs: the value, or the lowest; expression: the formula. */
  value1: string;
  /** cellIs, between and not between: the highest value. */
  value2: string;
  /** The text rules. */
  text: string;
  /** top10. */
  rank: string;
  percent: boolean;
  bottom: boolean;
  /** aboveAverage: "above" or "below". */
  average: string;
  equalAverage: boolean;
  stopIfTrue: boolean;
  /** One of the formats offered ("lightRed", "yellow", "green"), "custom" for the fields below, "none", or "keep": the format the rule has, as it is. */
  preset: string;
  useColor: boolean;
  /** "#RRGGBB". */
  color: string;
  useFill: boolean;
  fill: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  /** colorScale: "2" or "3" colours, from the lowest value to the highest (the second one is the midpoint of three). */
  scale: string;
  color1: string;
  color2: string;
  color3: string;
  /** dataBar. */
  barColor: string;
}

/** The formats the dialog offers, as Excel offers them. */
export const FORMAT_PRESETS: Record<string, { fill: string; color: string }> = {
  lightRed: { fill: '#FFC7CE', color: '#9C0006' },
  yellow: { fill: '#FFEB9C', color: '#9C5700' },
  green: { fill: '#C6EFCE', color: '#006100' },
};

/** Kinds of rule that give a format of their own choosing (the others colour by themselves). */
export const hasFormat = (type: string): boolean => type !== 'colorScale' && type !== 'dataBar';

const plainColor = (color: Color | undefined): string | undefined => (color && typeof color.rgb === 'string' && /^[0-9A-Fa-f]{6,8}$/.test(color.rgb) ? `#${color.rgb.slice(-6).toUpperCase()}` : undefined);

/** The kinds of rule that are stated by a text to look for. */
const TEXT_KINDS: Record<string, true> = { containsText: true, notContainsText: true, beginsWith: true, endsWith: true };

/** A stored value of a "cell value" rule the way it is typed: a number as the locale writes it, a text without its quotes, anything else as a formula. */
function shownCondition(text: string | undefined, locale: FormatLocale, date1904: boolean): string {
  if (typeof text !== 'string' || text === '') return '';
  const trimmed = text.trim();
  if (PLAIN_NUMBER.test(trimmed) && isFinite(Number(trimmed))) return cellEditText({ kind: 'number', value: Number(trimmed), style: 0 }, 'General', { locale, date1904 });
  if (trimmed.length >= 2 && trimmed.charAt(0) === '"' && trimmed.charAt(trimmed.length - 1) === '"') {
    const inner = trimmed.slice(1, -1);
    if (!/(^|[^"])"([^"]|$)/.test(inner.replace(/""/g, ''))) {
      const plain = inner.replace(/""/g, '"');
      // Only a text that is read back as that text when it is typed; "123" in quotes stays a formula.
      const parsed = plain === '' ? undefined : parseCellInput(plain, { locale, date1904 });
      if (parsed && parsed.kind === 'string' && parsed.value === plain && parsed.formula === undefined) return plain;
    }
  }
  return `=${formulaToLocal(text, locale)}`;
}

/**
 * The fields of the dialog for a rule (as conditionalRules() lists it); for no rule, the fields
 * a new rule of `type` starts with. `resolve` turns a colour of the file into "#RRGGBB" (theme
 * and indexed colours need the style sheet); without it only colours stated as such are shown.
 */
export function conditionalFields(info: ConditionalRuleInfo | undefined, locale: FormatLocale, date1904: boolean, type?: string, resolve?: (color: Color | undefined) => string | undefined): ConditionalFields {
  // As a colour field hands it back: "#RRGGBB" in capitals, so that a colour that was only looked at reads as unchanged.
  const upper = (text: string | undefined) => (text === undefined ? undefined : text.toUpperCase());
  const colorOf = (color: Color | undefined): string | undefined => upper((resolve ? resolve(color) : undefined) || plainColor(color));
  const fields: ConditionalFields = {
    type: typeof type === 'string' && (CONDITIONAL_TYPES as readonly string[]).indexOf(type) >= 0 ? type : 'cellIs',
    operator: 'greaterThan',
    value1: '',
    value2: '',
    text: '',
    rank: '10',
    percent: false,
    bottom: false,
    average: 'above',
    equalAverage: false,
    stopIfTrue: false,
    preset: 'lightRed',
    useColor: true,
    color: FORMAT_PRESETS.lightRed.color,
    useFill: true,
    fill: FORMAT_PRESETS.lightRed.fill,
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    scale: '3',
    color1: '#F8696B',
    color2: '#FFEB84',
    color3: '#63BE7B',
    barColor: '#638EC6',
  };
  if (!info) return fields;
  const rule = info.rule;
  fields.type = rule.type;
  // A text rule whose formula is not the one Excel writes for its text is the formula it is: stated by its text it would become another rule.
  const ownFormula = Object.prototype.hasOwnProperty.call(TEXT_KINDS, rule.type) && info.ranges.length > 0 && textRuleCell(rule, { row: info.ranges[0].top, col: info.ranges[0].left }) === undefined && typeof rule.formulas[0] === 'string';
  if (ownFormula) fields.type = 'expression';
  fields.stopIfTrue = rule.stopIfTrue === true;
  if (rule.type === 'cellIs') {
    fields.operator = typeof rule.operator === 'string' && (VALIDATION_OPERATORS as readonly string[]).indexOf(rule.operator) >= 0 ? rule.operator : 'between';
    fields.value1 = shownCondition(rule.formulas[0], locale, date1904);
    fields.value2 = shownCondition(rule.formulas[1], locale, date1904);
  } else if (rule.type === 'expression' || ownFormula) fields.value1 = typeof rule.formulas[0] === 'string' ? `=${formulaToLocal(rule.formulas[0], locale)}` : '';
  if (typeof rule.text === 'string') fields.text = rule.text;
  if (rule.type === 'top10') {
    fields.rank = String(rule.rank === undefined ? 10 : rule.rank);
    fields.percent = rule.percent === true;
    fields.bottom = rule.bottom === true;
  }
  if (rule.type === 'aboveAverage') {
    fields.average = rule.aboveAverage === false ? 'below' : 'above';
    fields.equalAverage = rule.equalAverage === true;
  }
  if (rule.colorScale) {
    const colors = rule.colorScale.colors.map(colorOf);
    fields.scale = colors.length >= 3 ? '3' : '2';
    if (colors[0]) fields.color1 = colors[0];
    if (colors.length >= 3 && colors[1]) fields.color2 = colors[1];
    const last = colors[colors.length >= 3 ? 2 : 1];
    if (last) fields.color3 = last;
  }
  if (rule.dataBar) fields.barColor = colorOf(rule.dataBar.color) || fields.barColor;

  // The format: none, one of those offered when it says exactly that and nothing else, or "the format the rule has" -
  // with what the fields can show of it, for whoever goes on to change it.
  const format = info.format;
  if (!format || Object.keys(format).length === 0) {
    fields.preset = 'none';
    return fields;
  }
  const shown = ruleFormatOf(format, resolve);
  if (shown.color !== undefined) shown.color = shown.color.toUpperCase();
  if (shown.fill !== undefined) shown.fill = shown.fill.toUpperCase();
  fields.preset = 'keep';
  const plainFont = !format.font || Object.keys(format.font).every((key) => key === 'color');
  const plain = plainFont && !format.border && !format.numFmt && (!format.fill || (format.fill.type === 'pattern' && format.fill.pattern === 'solid'));
  for (const name of Object.keys(FORMAT_PRESETS)) if (plain && FORMAT_PRESETS[name].fill === shown.fill && FORMAT_PRESETS[name].color === shown.color) fields.preset = name;
  fields.useColor = shown.color !== undefined;
  if (shown.color !== undefined) fields.color = shown.color;
  fields.useFill = shown.fill !== undefined;
  if (shown.fill !== undefined) fields.fill = shown.fill;
  fields.bold = shown.bold === true;
  fields.italic = shown.italic === true;
  fields.strike = shown.strike === true;
  fields.underline = shown.underline === true;
  return fields;
}

/** The format the fields ask for: one that is offered, the one typed in, none at all (null), or - undefined - the one the rule has. */
export function formatFromFields(fields: ConditionalFields): RuleFormat | null | undefined {
  const preset = Object.prototype.hasOwnProperty.call(FORMAT_PRESETS, fields.preset) ? FORMAT_PRESETS[fields.preset] : undefined;
  if (preset) return { fill: preset.fill, color: preset.color };
  if (fields.preset === 'none') return null;
  if (fields.preset !== 'custom') return undefined;
  const format: RuleFormat = {};
  if (fields.useColor) format.color = fields.color;
  if (fields.useFill) format.fill = fields.fill;
  if (fields.bold) format.bold = true;
  if (fields.italic) format.italic = true;
  if (fields.underline) format.underline = true;
  if (fields.strike) format.strike = true;
  return format;
}

const CONTENT_FIELDS = ['type', 'operator', 'value1', 'value2', 'text', 'rank', 'percent', 'bottom', 'average', 'equalAverage', 'scale', 'color1', 'color2', 'color3', 'barColor'] as const;
const FORMAT_FIELDS = ['preset', 'useColor', 'color', 'useFill', 'fill', 'bold', 'italic', 'underline', 'strike'] as const;

/**
 * What the dialog changes of the rule it was opened with, as updateConditionalRule() takes it:
 * only what a field says differently from when the dialog was opened. Nothing touched is an
 * empty patch, which is no change at all. A value that was not touched keeps what the rule
 * holds, as the file wrote it. Throws FieldError like conditionalFromFields(), and for cells
 * (`ranges`) that are none.
 */
export function conditionalPatch(
  fields: ConditionalFields,
  ranges: string,
  initial: { fields: ConditionalFields; ranges: string },
  info: ConditionalRuleInfo,
  locale: FormatLocale,
  strings: RuleStrings,
  options: FieldOptions = {},
): ConditionalPatch {
  const now = fields as unknown as Record<string, unknown>;
  const before = initial.fields as unknown as Record<string, unknown>;
  const patch: ConditionalPatch = {};
  if (ranges !== initial.ranges) {
    if (ranges.trim() === '') throw new FieldError(strings.needRange, 'ranges');
    const cells = parseRangeList(ranges, locale);
    if (!cells) throw new FieldError(strings.badRange, 'ranges');
    patch.ranges = cells;
  }
  const formatChanged = FORMAT_FIELDS.some((key) => now[key] !== before[key]);
  if (CONTENT_FIELDS.some((key) => now[key] !== before[key])) {
    const sameKind = fields.type === initial.fields.type;
    const keep1 = sameKind && fields.value1 === initial.fields.value1 && typeof info.rule.formulas[0] === 'string';
    const keep2 = sameKind && fields.value2 === initial.fields.value2 && typeof info.rule.formulas[1] === 'string';
    // A value that is kept is not read: any formula stands in for it while the others are.
    const settings = conditionalFromFields({ ...fields, value1: keep1 ? '=1' : fields.value1, value2: keep2 ? '=1' : fields.value2 }, locale, strings, options);
    if (settings.formulas) {
      if (keep1 && settings.formulas.length > 0) settings.formulas[0] = info.rule.formulas[0];
      if (keep2 && settings.formulas.length > 1) settings.formulas[1] = info.rule.formulas[1];
    }
    delete settings.stopIfTrue;
    if (!formatChanged) delete settings.format;
    patch.rule = settings;
  } else if (formatChanged && hasFormat(fields.type)) {
    const format = formatFromFields(fields);
    if (format !== undefined) patch.format = format;
  }
  if (fields.stopIfTrue !== initial.fields.stopIfTrue) patch.stopIfTrue = fields.stopIfTrue === true;
  return patch;
}

/**
 * The rule the fields of the dialog stand for, as addConditionalRule() takes it. Throws
 * FieldError (in the language of `strings`) for a field that cannot be taken.
 */
export function conditionalFromFields(fields: ConditionalFields, locale: FormatLocale, strings: RuleStrings, options: FieldOptions = {}): ConditionalSettings {
  const type = ((CONDITIONAL_TYPES as readonly string[]).indexOf(fields.type) >= 0 ? fields.type : 'cellIs') as ConditionalType;
  const date1904 = options.date1904 === true;
  const settings: ConditionalSettings = { type, stopIfTrue: fields.stopIfTrue === true };
  const formula = (text: string, field: string, label: string): string => {
    let stored: string;
    try {
      stored = formulaFromLocal(text, locale, { isName: options.isName });
      if (stored.trim() !== '') parseFormula(stored);
    } catch (error) {
      throw new FieldError(strings.badRuleFormula(label), field);
    }
    if (stored.trim() === '') throw new FieldError(strings.needValue(label), field);
    return stored;
  };
  /** A value a cell is compared with, as a file stores it: a number, a text in quotes, a formula. */
  const condition = (typed: string, field: string, label: string): string => {
    const text = typed.trim();
    if (text === '') throw new FieldError(strings.needValue(label), field);
    if (text.charAt(0) === '=') return formula(text.slice(1), field, label);
    const parsed = parseCellInput(text, { locale, date1904 });
    if (parsed.kind === 'number' && typeof parsed.value === 'number' && isFinite(parsed.value)) return String(parsed.value);
    if (parsed.kind === 'boolean') return parsed.value ? 'TRUE' : 'FALSE';
    if (parsed.kind === 'error' && typeof parsed.value === 'string') return parsed.value;
    return `"${text.replace(/"/g, '""')}"`;
  };
  switch (type) {
    case 'cellIs': {
      const operator = ((VALIDATION_OPERATORS as readonly string[]).indexOf(fields.operator) >= 0 ? fields.operator : 'between') as ValidationOperator;
      const two = operator === 'between' || operator === 'notBetween';
      settings.operator = operator;
      settings.formulas = [condition(fields.value1, 'value1', two ? strings.minimum : strings.value)];
      if (two) settings.formulas.push(condition(fields.value2, 'value2', strings.maximum));
      break;
    }
    case 'expression': {
      const text = fields.value1.trim();
      settings.formulas = [formula(text.charAt(0) === '=' ? text.slice(1) : text, 'value1', strings.formula)];
      break;
    }
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith':
      if (fields.text === '') throw new FieldError(strings.needValue(strings.text), 'text');
      if (fields.text.length > MAX_CONDITIONAL_TEXT) throw new FieldError(strings.textTooLong(strings.text, MAX_CONDITIONAL_TEXT), 'text');
      settings.text = fields.text;
      break;
    case 'top10': {
      const percent = fields.percent === true;
      const limit = percent ? 100 : 1000;
      const rank = /^\d{1,4}$/.test(fields.rank.trim()) ? Number(fields.rank.trim()) : NaN;
      if (!(rank >= 1 && rank <= limit)) throw new FieldError(strings.needWholeUpTo(strings.rank, limit), 'rank');
      settings.rank = rank;
      settings.percent = percent;
      settings.bottom = fields.bottom === true;
      break;
    }
    case 'aboveAverage':
      settings.aboveAverage = fields.average !== 'below';
      settings.equalAverage = fields.equalAverage === true;
      break;
    case 'colorScale':
      settings.colors = fields.scale === '2' ? [fields.color1, fields.color3] : [fields.color1, fields.color2, fields.color3];
      break;
    case 'dataBar':
      settings.colors = [fields.barColor];
      break;
    default:
    // The other kinds say everything by being what they are.
  }
  if (hasFormat(type)) {
    const format = formatFromFields(fields);
    if (format !== undefined) settings.format = format;
  }
  return settings;
}

/** What a rule says, in one line, for the list of rules. */
export function describeRule(rule: ConditionalRule, strings: RuleStrings, locale: FormatLocale, date1904: boolean): string {
  const named = (table: Record<string, string>, key: string | undefined, fallback: string) => (key !== undefined && Object.prototype.hasOwnProperty.call(table, key) ? table[key] : fallback);
  switch (rule.type) {
    case 'cellIs': {
      const operator = named(strings.operators, rule.operator, rule.operator || '');
      const first = shownCondition(rule.formulas[0], locale, date1904);
      const two = rule.operator === 'between' || rule.operator === 'notBetween';
      return `${strings.cellValue} ${operator} ${first}${two ? ` ${strings.and} ${shownCondition(rule.formulas[1], locale, date1904)}` : ''}`;
    }
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith':
      return `${strings.textRules[rule.type]} "${typeof rule.text === 'string' ? rule.text : ''}"`;
    case 'top10':
      return strings.topItems(rule.rank === undefined ? 10 : rule.rank, rule.percent === true, rule.bottom === true);
    case 'aboveAverage':
      return strings.average(rule.aboveAverage !== false, rule.equalAverage === true);
    case 'expression':
      return `${strings.formulaIs}: =${formulaToLocal(typeof rule.formulas[0] === 'string' ? rule.formulas[0] : '', locale)}`;
    case 'timePeriod':
      return `${strings.datesOccurring}: ${named(strings.periods, rule.timePeriod, rule.timePeriod || '')}`;
    default:
      return named(strings.conditionalTypes, rule.type, rule.type);
  }
}

/**
 * The ranges a list typed into a field names: cells ("A1"), ranges ("A1:B10"), whole columns
 * ("A:C") and whole rows ("2:5"), with or without "$", separated by ";", "," or spaces.
 * Undefined for a text that is anything else, or names nothing.
 */
export function parseRangeList(text: string, _locale?: FormatLocale): CellRange[] | undefined {
  if (typeof text !== 'string' || text.length > 32768) return undefined;
  const out: CellRange[] = [];
  for (const item of text.split(/[;,\s]+/)) {
    if (item === '') continue;
    const ref = item.replace(/\$/g, '');
    let range = parseRange(ref);
    if (!range) {
      const cols = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(ref);
      const rows = /^(\d{1,7}):(\d{1,7})$/.exec(ref);
      if (cols) {
        const a = parseCellRef(`${cols[1]}1`);
        const b = parseCellRef(`${cols[2]}1`);
        if (a && b) range = { top: 0, left: Math.min(a.col, b.col), bottom: MAX_ROWS - 1, right: Math.max(a.col, b.col) };
      } else if (rows) {
        const a = Number(rows[1]);
        const b = Number(rows[2]);
        if (a >= 1 && b >= 1 && a <= MAX_ROWS && b <= MAX_ROWS) range = { top: Math.min(a, b) - 1, left: 0, bottom: Math.max(a, b) - 1, right: MAX_COLS - 1 };
      }
    }
    if (!range) return undefined;
    out.push(range);
  }
  return out.length > 0 ? out : undefined;
}

/** Ranges as they are typed into that field. */
export function rangeListText(ranges: readonly CellRange[], locale: FormatLocale): string {
  return ranges.map(rangeName).join(`${listSeparator(locale)} `);
}

// ---------------------------------------------------------------------------
// The dialog

let nextId = 1;

export interface DialogOptions {
  /** Positioned element the dialog covers. */
  parent: HTMLElement;
  title: string;
  /** Marks the dialog for the page and its tests: data-bym-dialog. */
  name: string;
  /** Runs once when the dialog is gone, whichever way it went. */
  onClose(): void;
  /** Opened by a finger: fields and buttons are larger. */
  coarse?: boolean;
}

const stop = (event: Event): void => event.stopPropagation();

/** Events of the dialog that the view around it must not take for its own. */
const CONTAINED = ['keyup', 'keypress', 'mousedown', 'mouseup', 'click', 'dblclick', 'pointerdown', 'pointerup', 'touchstart', 'touchend', 'wheel', 'contextmenu', 'copy', 'cut', 'paste', 'input', 'compositionstart', 'compositionend'];

/**
 * A modal dialog: a title, a body the caller fills, a line for what is wrong, and buttons.
 * Tab stays inside, Escape closes, Enter runs the primary button (in a field of several
 * lines Ctrl+Enter does, Enter is a line break).
 */
export class Dialog {
  readonly element: HTMLDivElement;
  readonly panel: HTMLDivElement;
  readonly body: HTMLDivElement;
  private readonly alert: HTMLDivElement;
  private readonly actions: HTMLDivElement;
  private readonly onClose: () => void;
  private readonly coarse: boolean;
  private primary: (() => void) | undefined;
  private invalid: HTMLElement | undefined;
  private closed = false;
  private pristine: string | undefined;

  constructor(options: DialogOptions) {
    this.onClose = options.onClose;
    this.coarse = options.coarse === true;
    const element = document.createElement('div');
    element.className = 'bym-dialog-backdrop';
    // What the layout depends on is set inline: it holds where a style element is not allowed.
    element.style.cssText = 'position:absolute;left:0;top:0;right:0;bottom:0;z-index:30;display:flex;align-items:center;justify-content:center;box-sizing:border-box;padding:8px;overflow:hidden';
    const panel = document.createElement('div');
    panel.className = 'bym-dialog';
    panel.setAttribute('role', 'dialog');
    panel.setAttribute('aria-modal', 'true');
    panel.setAttribute('data-bym-dialog', options.name);
    panel.style.cssText = 'display:flex;flex-direction:column;box-sizing:border-box;width:380px;max-width:100%;max-height:100%;min-width:0;min-height:0;overflow:hidden';
    const heading = document.createElement('div');
    heading.className = 'bym-dialog-title';
    heading.id = `bym-dialog-${nextId++}`;
    heading.style.cssText = 'flex:none';
    heading.textContent = options.title;
    panel.setAttribute('aria-labelledby', heading.id);
    const body = document.createElement('div');
    body.className = 'bym-dialog-body';
    body.style.cssText = 'flex:1 1 auto;min-height:0;overflow-x:hidden;overflow-y:auto;overscroll-behavior:contain';
    const alert = document.createElement('div');
    alert.className = 'bym-dialog-alert';
    alert.setAttribute('role', 'alert');
    alert.style.cssText = 'flex:none';
    const actions = document.createElement('div');
    actions.className = 'bym-dialog-actions';
    actions.style.cssText = 'display:flex;flex:none;flex-wrap:wrap;justify-content:flex-end;align-items:center';
    panel.appendChild(heading);
    panel.appendChild(body);
    panel.appendChild(alert);
    panel.appendChild(actions);
    element.appendChild(panel);
    this.element = element;
    this.panel = panel;
    this.body = body;
    this.alert = alert;
    this.actions = actions;
    element.addEventListener('keydown', this.onKeyDown);
    element.addEventListener('mousedown', this.onMouseDown);
    for (const type of CONTAINED) element.addEventListener(type, stop);
    options.parent.appendChild(element);
  }

  get isOpen(): boolean {
    return !this.closed;
  }

  /** Adds a button. `primary`: the one Enter runs; `start`: stands apart from the others, at the start of the row. */
  addAction(label: string, run: () => void, kind: 'primary' | 'plain' | 'start' = 'plain'): HTMLButtonElement {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = kind === 'primary' ? 'bym-dialog-button bym-dialog-primary' : 'bym-dialog-button';
    button.textContent = label;
    if (kind === 'start') button.style.marginRight = 'auto';
    button.addEventListener('click', () => {
      if (!this.closed) run();
    });
    if (kind === 'primary') this.primary = run;
    this.actions.appendChild(button);
    return button;
  }

  /** A row with a label and its control; the label names the control for assistive technology. */
  addField(parent: HTMLElement, label: string, control: HTMLElement, name: string): HTMLDivElement {
    const row = document.createElement('div');
    row.className = 'bym-dialog-row';
    const id = `bym-dialog-${nextId++}`;
    control.id = id;
    control.setAttribute('data-bym-field', name);
    const text = document.createElement('label');
    text.className = 'bym-dialog-label';
    text.htmlFor = id;
    text.textContent = label;
    row.appendChild(text);
    row.appendChild(control);
    parent.appendChild(row);
    return row;
  }

  /** A check box with its label behind it. */
  addCheck(parent: HTMLElement, label: string, name: string): HTMLInputElement {
    const row = document.createElement('label');
    row.className = 'bym-dialog-check';
    const box = document.createElement('input');
    box.type = 'checkbox';
    box.className = 'bym-dialog-box';
    box.setAttribute('data-bym-field', name);
    const text = document.createElement('span');
    text.textContent = label;
    row.appendChild(box);
    row.appendChild(text);
    parent.appendChild(row);
    return box;
  }

  input(maxLength?: number): HTMLInputElement {
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'bym-dialog-input';
    input.setAttribute('autocomplete', 'off');
    input.setAttribute('autocapitalize', 'off');
    input.setAttribute('autocorrect', 'off');
    input.spellcheck = false;
    if (maxLength !== undefined) input.maxLength = maxLength;
    // A field with smaller text makes a phone zoom into the page.
    if (this.coarse) input.style.fontSize = '16px';
    return input;
  }

  /** A colour field; its value is "#rrggbb". */
  color(): HTMLInputElement {
    const input = document.createElement('input');
    input.type = 'color';
    input.className = 'bym-dialog-color';
    return input;
  }

  textarea(maxLength?: number): HTMLTextAreaElement {
    const input = document.createElement('textarea');
    input.className = 'bym-dialog-input bym-dialog-text';
    input.rows = 2;
    if (maxLength !== undefined) input.maxLength = maxLength;
    if (this.coarse) input.style.fontSize = '16px';
    return input;
  }

  select(items: [value: string, label: string][]): HTMLSelectElement {
    const select = document.createElement('select');
    select.className = 'bym-dialog-input bym-dialog-select';
    for (const [value, label] of items) {
      const option = document.createElement('option');
      option.value = value;
      option.textContent = label;
      select.appendChild(option);
    }
    if (this.coarse) select.style.fontSize = '16px';
    return select;
  }

  section(title: string): HTMLFieldSetElement {
    const set = document.createElement('fieldset');
    set.className = 'bym-dialog-section';
    const legend = document.createElement('legend');
    legend.className = 'bym-dialog-legend';
    legend.textContent = title;
    set.appendChild(legend);
    this.body.appendChild(set);
    return set;
  }

  /** Says what is wrong, and points at the field it is about; without a message, takes back what was said. */
  showError(message: string, field?: HTMLElement): void {
    if (this.invalid) this.invalid.removeAttribute('aria-invalid');
    this.invalid = undefined;
    this.alert.textContent = message;
    if (!field || message === '') return;
    field.setAttribute('aria-invalid', 'true');
    this.invalid = field;
    try {
      field.focus();
      if (typeof field.scrollIntoView === 'function') field.scrollIntoView({ block: 'nearest' });
    } catch {
      // Not rendered: the message says it all the same.
    }
  }

  /** Puts the focus on the first control of the dialog. */
  focusFirst(): void {
    const first = this.focusable()[0];
    if (first) first.focus({ preventScroll: true });
  }

  close(): void {
    if (this.closed) return;
    this.closed = true;
    this.element.removeEventListener('keydown', this.onKeyDown);
    this.element.removeEventListener('mousedown', this.onMouseDown);
    for (const type of CONTAINED) this.element.removeEventListener(type, stop);
    if (this.element.parentNode) this.element.parentNode.removeChild(this.element);
    this.onClose();
  }

  /** Runs the primary button, as Enter does. */
  submit(): void {
    if (!this.closed && this.primary) this.primary();
  }

  /** What every field of the dialog holds, as one text: equal for two moments exactly when no field differs. */
  private fieldState(): string {
    const fields = this.panel.querySelectorAll<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>('input, select, textarea');
    let out = '';
    for (let i = 0; i < fields.length; i++) {
      const field = fields[i];
      const checked = field instanceof HTMLInputElement && (field.type === 'checkbox' || field.type === 'radio') ? (field.checked ? '1' : '0') : '';
      out += `${field.value.length}:${field.value}${checked}\u0001`;
    }
    return out;
  }

  /** Remembers what the fields hold now as what the dialog was opened with (see `changed`). */
  markPristine(): void {
    this.pristine = this.fieldState();
  }

  /**
   * Whether a field of the dialog holds something else than when it was opened. A dialog
   * with OK and Cancel in which nothing was changed can be closed by whoever needs the
   * workbook as it stands; one with changes is the user's to confirm or to give up.
   */
  get changed(): boolean {
    return !this.closed && this.pristine !== undefined && this.fieldState() !== this.pristine;
  }

  private focusable(): HTMLElement[] {
    const found = this.panel.querySelectorAll<HTMLElement>('button, input, select, textarea, [tabindex]');
    const out: HTMLElement[] = [];
    for (let i = 0; i < found.length; i++) {
      const el = found[i];
      if ((el as HTMLButtonElement).disabled || el.tabIndex < 0) continue;
      // A control of a part of the dialog that is not shown takes no focus.
      if (el.getClientRects().length === 0) continue;
      out.push(el);
    }
    return out;
  }

  private onMouseDown = (event: MouseEvent): void => {
    // A press beside the dialog leaves the focus where it is: the dialog is all there is to work with.
    if (event.target === this.element) event.preventDefault();
  };

  private onKeyDown = (event: KeyboardEvent): void => {
    // Whatever the key is, it is the dialog's: the grid behind it does not see it.
    event.stopPropagation();
    if (this.closed || event.isComposing || event.keyCode === 229) return;
    if (event.key === 'Escape') {
      event.preventDefault();
      this.close();
      return;
    }
    if (event.key === 'Tab') {
      const items = this.focusable();
      if (items.length === 0) {
        event.preventDefault();
        return;
      }
      const root = this.panel.getRootNode() as Document | ShadowRoot;
      const at = items.indexOf(root.activeElement as HTMLElement);
      const next = event.shiftKey ? (at <= 0 ? items.length - 1 : at - 1) : at < 0 || at === items.length - 1 ? 0 : at + 1;
      event.preventDefault();
      items[next].focus();
      return;
    }
    if (event.key !== 'Enter' || event.altKey) return;
    const target = event.target as HTMLElement | null;
    const tag = target ? target.tagName : '';
    // A button does what it says; a field of several lines takes Enter as a line break.
    if (tag === 'BUTTON') return;
    if (tag === 'TEXTAREA' && !(event.ctrlKey || event.metaKey)) return;
    event.preventDefault();
    this.submit();
  };
}

// ---------------------------------------------------------------------------
// Data validation

export interface ValidationDialogOptions {
  parent: HTMLElement;
  strings: RuleStrings;
  locale: FormatLocale;
  fields: ValidationFields;
  /** The selection has cells of several rules, or cells with a rule beside cells without. */
  mixed: boolean;
  coarse?: boolean;
  /** Applies the rule; a FieldError it throws is shown at its field, any other error as a message, and the dialog stays open. */
  onApply(fields: ValidationFields): void;
  /** Takes the rules off the selection. */
  onClear(): void;
  /** What an error that is no FieldError says, in the language of the dialog. */
  explain(error: unknown): string;
  onClose(): void;
}

/** The dialog that sets the data validation of the selection. */
export class ValidationDialog {
  readonly dialog: Dialog;
  private readonly strings: RuleStrings;
  private readonly controls: {
    type: HTMLSelectElement;
    operator: HTMLSelectElement;
    value1: HTMLInputElement;
    value2: HTMLInputElement;
    allowBlank: HTMLInputElement;
    inCellDropdown: HTMLInputElement;
    showInputMessage: HTMLInputElement;
    promptTitle: HTMLInputElement;
    prompt: HTMLTextAreaElement;
    showErrorMessage: HTMLInputElement;
    errorStyle: HTMLSelectElement;
    errorTitle: HTMLInputElement;
    error: HTMLTextAreaElement;
  };
  private readonly rows: { operator: HTMLElement; value1: HTMLElement; value2: HTMLElement; hint: HTMLElement; dropdown: HTMLElement; blank: HTMLElement };
  private readonly separator: string;

  constructor(options: ValidationDialogOptions) {
    const strings = options.strings;
    this.strings = strings;
    this.separator = listSeparator(options.locale);
    const dialog = new Dialog({ parent: options.parent, title: strings.validationTitle, name: 'validation', onClose: options.onClose, coarse: options.coarse });
    this.dialog = dialog;

    if (options.mixed) {
      const note = document.createElement('div');
      note.className = 'bym-dialog-note';
      note.setAttribute('data-bym-field', 'mixed');
      note.textContent = strings.mixedSelection;
      dialog.body.appendChild(note);
    }

    const settings = dialog.section(strings.settings);
    const type = dialog.select(VALIDATION_TYPES.map((name) => [name, strings.types[name]] as [string, string]));
    dialog.addField(settings, strings.allow, type, 'type');
    const operator = dialog.select(VALIDATION_OPERATORS.map((name) => [name, strings.operators[name]] as [string, string]));
    const operatorRow = dialog.addField(settings, strings.data, operator, 'operator');
    const value1 = dialog.input();
    const value1Row = dialog.addField(settings, strings.value, value1, 'value1');
    const hint = document.createElement('div');
    hint.className = 'bym-dialog-hint';
    hint.textContent = strings.sourceHint(this.separator);
    hint.id = `bym-dialog-${nextId++}`;
    settings.appendChild(hint);
    const value2 = dialog.input();
    const value2Row = dialog.addField(settings, strings.maximum, value2, 'value2');
    const allowBlank = dialog.addCheck(settings, strings.ignoreBlank, 'allowBlank');
    const inCellDropdown = dialog.addCheck(settings, strings.inCellDropdown, 'inCellDropdown');

    const input = dialog.section(strings.inputMessage);
    const showInputMessage = dialog.addCheck(input, strings.showInputMessage, 'showInputMessage');
    const promptTitle = dialog.input();
    dialog.addField(input, strings.title, promptTitle, 'promptTitle');
    const prompt = dialog.textarea();
    dialog.addField(input, strings.message, prompt, 'prompt');

    const alert = dialog.section(strings.errorAlert);
    const showErrorMessage = dialog.addCheck(alert, strings.showErrorAlert, 'showErrorMessage');
    const errorStyle = dialog.select(['stop', 'warning', 'information'].map((name) => [name, strings.styles[name]] as [string, string]));
    dialog.addField(alert, strings.style, errorStyle, 'errorStyle');
    const errorTitle = dialog.input();
    dialog.addField(alert, strings.title, errorTitle, 'errorTitle');
    const error = dialog.textarea();
    dialog.addField(alert, strings.message, error, 'error');

    this.controls = { type, operator, value1, value2, allowBlank, inCellDropdown, showInputMessage, promptTitle, prompt, showErrorMessage, errorStyle, errorTitle, error };
    this.rows = { operator: operatorRow, value1: value1Row, value2: value2Row, hint, dropdown: inCellDropdown.parentElement as HTMLElement, blank: allowBlank.parentElement as HTMLElement };

    // What the file says is only ever the value of a field.
    const fields = options.fields;
    type.value = fields.type;
    operator.value = fields.operator;
    value1.value = fields.value1;
    value2.value = fields.value2;
    allowBlank.checked = fields.allowBlank;
    inCellDropdown.checked = fields.inCellDropdown;
    showInputMessage.checked = fields.showInputMessage;
    promptTitle.value = fields.promptTitle;
    prompt.value = fields.prompt;
    showErrorMessage.checked = fields.showErrorMessage;
    errorStyle.value = fields.errorStyle;
    errorTitle.value = fields.errorTitle;
    error.value = fields.error;
    type.addEventListener('change', () => this.arrange());
    operator.addEventListener('change', () => this.arrange());
    this.arrange();

    const attempt = (run: () => void) => {
      dialog.showError('');
      try {
        run();
      } catch (problem) {
        if (problem instanceof FieldError) dialog.showError(problem.message, (this.controls as Record<string, HTMLElement>)[problem.field]);
        else dialog.showError(options.explain(problem));
        return;
      }
      dialog.close();
    };
    dialog.addAction(strings.clearAll, () => attempt(() => options.onClear()), 'start').setAttribute('data-bym-action', 'clear');
    dialog.addAction(strings.cancel, () => dialog.close()).setAttribute('data-bym-action', 'cancel');
    dialog.addAction(strings.ok, () => attempt(() => options.onApply(this.fields())), 'primary').setAttribute('data-bym-action', 'ok');
    dialog.focusFirst();
    dialog.markPristine();
  }

  /** What the fields hold now. */
  fields(): ValidationFields {
    const c = this.controls;
    return {
      type: c.type.value,
      operator: c.operator.value,
      value1: c.value1.value,
      value2: c.value2.value,
      allowBlank: c.allowBlank.checked,
      inCellDropdown: c.inCellDropdown.checked,
      showInputMessage: c.showInputMessage.checked,
      promptTitle: c.promptTitle.value,
      // A text field turns every kind of line break into "\n".
      prompt: c.prompt.value,
      showErrorMessage: c.showErrorMessage.checked,
      errorStyle: c.errorStyle.value,
      errorTitle: c.errorTitle.value,
      error: c.error.value,
    };
  }

  /** Shows the fields the kind of rule has, under the names it gives them. */
  private arrange(): void {
    const type = this.controls.type.value;
    const operator = this.controls.operator.value;
    const show = (el: HTMLElement, shown: boolean) => {
      el.style.display = shown ? '' : 'none';
    };
    show(this.rows.operator, hasOperator(type));
    show(this.rows.value1, type !== 'none');
    show(this.rows.value2, hasSecondValue(type, operator));
    show(this.rows.hint, type === 'list');
    show(this.rows.dropdown, type === 'list');
    show(this.rows.blank, type !== 'none');
    const label = (row: HTMLElement, text: string) => {
      const found = row.getElementsByTagName('label')[0];
      if (found) found.textContent = text;
    };
    label(this.rows.value1, valueLabel(type, operator, false, this.strings));
    label(this.rows.value2, valueLabel(type, operator, true, this.strings));
    if (type === 'list') this.controls.value1.setAttribute('aria-describedby', this.rows.hint.id);
    else this.controls.value1.removeAttribute('aria-describedby');
  }
}

// ---------------------------------------------------------------------------
// One conditional rule

export interface ConditionalDialogOptions {
  parent: HTMLElement;
  strings: RuleStrings;
  locale: FormatLocale;
  fields: ConditionalFields;
  /** The cells the rule applies to, as they are typed (see rangeListText()). */
  ranges: string;
  /** The field that gets the focus first; the first control of the dialog without it. */
  focus?: string;
  /** The rule has a format: "the format it has" is one of the formats to choose from. */
  keepFormat?: boolean;
  coarse?: boolean;
  /** Applies the rule; a FieldError it throws is shown at its field, any other error as a message, and the dialog stays open. */
  onApply(fields: ConditionalFields, ranges: string): void;
  explain(error: unknown): string;
  onClose(): void;
}

/** The field a kind of rule is typed into first. */
function firstFieldOf(type: string): string {
  if (type === 'cellIs' || type === 'expression') return 'value1';
  if (type === 'containsText' || type === 'notContainsText' || type === 'beginsWith' || type === 'endsWith') return 'text';
  if (type === 'top10') return 'rank';
  if (type === 'aboveAverage') return 'average';
  if (type === 'colorScale') return 'scale';
  if (type === 'dataBar') return 'barColor';
  return 'preset';
}

/** The dialog that states one conditional rule: what it looks for, where, and the format it gives. */
export class ConditionalDialog {
  readonly dialog: Dialog;
  private readonly strings: RuleStrings;
  private readonly controls: Record<string, HTMLInputElement | HTMLSelectElement>;
  private readonly rows: Record<string, HTMLElement>;

  constructor(options: ConditionalDialogOptions) {
    const strings = options.strings;
    this.strings = strings;
    const dialog = new Dialog({ parent: options.parent, title: strings.conditionalTitle, name: 'conditional', onClose: options.onClose, coarse: options.coarse });
    this.dialog = dialog;
    const controls: Record<string, HTMLInputElement | HTMLSelectElement> = {};
    const rows: Record<string, HTMLElement> = {};
    const body = dialog.body;
    const text = (name: string, label: string, parent: HTMLElement = body) => {
      const input = dialog.input();
      controls[name] = input;
      rows[name] = dialog.addField(parent, label, input, name);
    };
    const list = (name: string, label: string, items: [string, string][], parent: HTMLElement = body) => {
      const select = dialog.select(items);
      controls[name] = select;
      rows[name] = dialog.addField(parent, label, select, name);
    };
    const check = (name: string, label: string, parent: HTMLElement = body) => {
      const box = dialog.addCheck(parent, label, name);
      controls[name] = box;
      rows[name] = box.parentElement as HTMLElement;
    };
    const colour = (name: string, label: string, parent: HTMLElement = body) => {
      const input = dialog.color();
      controls[name] = input;
      rows[name] = dialog.addField(parent, label, input, name);
    };

    text('ranges', strings.appliesTo);
    list('type', strings.ruleType, CONDITIONAL_TYPES.map((name) => [name, strings.conditionalTypes[name]] as [string, string]));
    list('operator', strings.data, VALIDATION_OPERATORS.map((name) => [name, strings.operators[name]] as [string, string]));
    text('value1', strings.value);
    text('value2', strings.maximum);
    text('text', strings.text);
    text('rank', strings.rank);
    check('percent', strings.percentOfRange);
    check('bottom', strings.lowestValues);
    list('average', strings.averageSide, [['above', strings.averages.above], ['below', strings.averages.below]]);
    check('equalAverage', strings.includeAverage);
    list('scale', strings.scaleKind, [['2', strings.scales['2']], ['3', strings.scales['3']]]);
    colour('color1', strings.lowest);
    colour('color2', strings.midpoint);
    colour('color3', strings.highest);
    colour('barColor', strings.barColor);

    const format = dialog.section(strings.format);
    rows.format = format;
    list('preset', strings.formatWith, (options.keepFormat ? ['keep'] : []).concat(['lightRed', 'yellow', 'green', 'custom', 'none']).map((name) => [name, strings.presets[name]] as [string, string]), format);
    check('useColor', strings.textColor, format);
    colour('color', strings.textColor, format);
    check('useFill', strings.fillColor, format);
    colour('fill', strings.fillColor, format);
    check('bold', strings.bold, format);
    check('italic', strings.italic, format);
    check('underline', strings.underline, format);
    check('strike', strings.strike, format);
    check('stopIfTrue', strings.stopIfTrue);

    this.controls = controls;
    this.rows = rows;
    // What the file says is only ever the value of a field.
    const fields = options.fields as unknown as Record<string, string | boolean>;
    for (const name of Object.keys(controls)) {
      const control = controls[name];
      if (name === 'ranges') control.value = options.ranges;
      else if (control instanceof HTMLInputElement && control.type === 'checkbox') control.checked = fields[name] === true;
      else if (control instanceof HTMLInputElement && control.type === 'color') control.value = String(fields[name]).toLowerCase();
      else control.value = String(fields[name]);
    }
    for (const name of ['type', 'operator', 'preset', 'scale', 'useColor', 'useFill']) controls[name].addEventListener('change', () => this.arrange());
    this.arrange();

    dialog.addAction(strings.cancel, () => dialog.close()).setAttribute('data-bym-action', 'cancel');
    dialog
      .addAction(
        strings.ok,
        () => {
          dialog.showError('');
          try {
            options.onApply(this.fields(), controls.ranges.value);
          } catch (problem) {
            if (problem instanceof FieldError) dialog.showError(problem.message, controls[problem.field]);
            else dialog.showError(options.explain(problem));
            return;
          }
          dialog.close();
        },
        'primary',
      )
      .setAttribute('data-bym-action', 'ok');
    const first = controls[options.focus !== undefined ? options.focus : firstFieldOf(options.fields.type)];
    if (first && first.getClientRects().length > 0) first.focus({ preventScroll: true });
    else dialog.focusFirst();
    dialog.markPristine();
  }

  /** What the fields hold now. */
  fields(): ConditionalFields {
    const out: Record<string, string | boolean> = {};
    for (const name of Object.keys(this.controls)) {
      if (name === 'ranges') continue;
      const control = this.controls[name];
      if (control instanceof HTMLInputElement && control.type === 'checkbox') out[name] = control.checked;
      else if (control instanceof HTMLInputElement && control.type === 'color') out[name] = control.value.toUpperCase();
      else out[name] = control.value;
    }
    return out as unknown as ConditionalFields;
  }

  /** Shows the fields the kind of rule has, under the names it gives them. */
  private arrange(): void {
    const c = this.controls;
    const type = c.type.value;
    const operator = c.operator.value;
    const two = type === 'cellIs' && (operator === 'between' || operator === 'notBetween');
    const custom = hasFormat(type) && c.preset.value === 'custom';
    const textRule = type === 'containsText' || type === 'notContainsText' || type === 'beginsWith' || type === 'endsWith';
    const shown: Record<string, boolean> = {
      operator: type === 'cellIs',
      value1: type === 'cellIs' || type === 'expression',
      value2: two,
      text: textRule,
      rank: type === 'top10',
      percent: type === 'top10',
      bottom: type === 'top10',
      average: type === 'aboveAverage',
      equalAverage: type === 'aboveAverage',
      scale: type === 'colorScale',
      color1: type === 'colorScale',
      color2: type === 'colorScale' && c.scale.value === '3',
      color3: type === 'colorScale',
      barColor: type === 'dataBar',
      format: hasFormat(type),
      useColor: custom,
      color: custom && (c.useColor as HTMLInputElement).checked,
      useFill: custom,
      fill: custom && (c.useFill as HTMLInputElement).checked,
      bold: custom,
      italic: custom,
      underline: custom,
      strike: custom,
    };
    for (const name of Object.keys(shown)) this.rows[name].style.display = shown[name] ? '' : 'none';
    const label = this.rows.value1.getElementsByTagName('label')[0];
    if (label) label.textContent = type === 'expression' ? this.strings.formula : two ? this.strings.minimum : this.strings.value;
  }
}

// ---------------------------------------------------------------------------
// The list of conditional rules

/** Most rules the list shows; a sheet may have tens of thousands, and every change lists them again. */
export const MAX_MANAGER_ROWS = 200;

/** One rule as the list shows it. */
export interface ManagerRow {
  /** What the commands name the rule by. */
  index: number;
  description: string;
  /** The cells it applies to, as they are typed. */
  ranges: string;
  stopIfTrue: boolean;
  /** What it says can be changed (Edit). */
  editable: boolean;
  removable: boolean;
  movable: boolean;
  /** Where it applies and whether it stops the others can be changed. */
  adjustable: boolean;
  /** Why something of it cannot be done, for the controls that are off. */
  reason?: string;
}

export interface ConditionalManagerOptions {
  parent: HTMLElement;
  strings: RuleStrings;
  scope: 'selection' | 'sheet';
  coarse?: boolean;
  /**
   * The rules to show, in the order they are applied in: those from place `from` on (no more
   * than MAX_MANAGER_ROWS of them), how many there are in all, and the rule right before and
   * right after the ones shown (what the first can be moved above and the last below); asked
   * again after every change.
   */
  rows(scope: 'selection' | 'sheet', from: number): { rows: ManagerRow[]; total: number; before?: ManagerRow; after?: ManagerRow };
  /** Each changes the workbook at once, as one undo step; an error is shown in the dialog. */
  onMove(index: number, to: number): void;
  onRanges(index: number, ranges: string): void;
  onStop(index: number, stop: boolean): void;
  onDelete(index: number): void;
  /** Open the dialog of a rule; the list is closed first and opened again afterwards by the caller. */
  onEdit(index: number, scope: 'selection' | 'sheet'): void;
  onNew(scope: 'selection' | 'sheet'): void;
  explain(error: unknown): string;
  onClose(): void;
}

/**
 * Which of `total` rules a list shows that begins at place `from` and shows `size` at a time:
 * the place it really begins at (a multiple of `size`, and one a rule is at unless there is
 * none), where it ends, and where the lists before and after it begin (undefined: there is none).
 */
export function managerPage(total: number, from: number, size: number = MAX_MANAGER_ROWS): { from: number; to: number; previous?: number; next?: number } {
  const count = total > 0 && total === Math.floor(total) ? total : 0;
  const step = size >= 1 ? Math.floor(size) : 1;
  let start = from > 0 && from === from ? Math.floor(from / step) * step : 0;
  if (start >= count) start = count > 0 ? Math.floor((count - 1) / step) * step : 0;
  const out: { from: number; to: number; previous?: number; next?: number } = { from: start, to: Math.min(start + step, count) };
  if (start > 0) out.previous = start - step;
  if (out.to < count) out.next = out.to;
  return out;
}

/** The list of the conditional rules of the selection or the sheet: their order, their cells, and the way to each rule's own dialog. */
export class ConditionalManager {
  readonly dialog: Dialog;
  private readonly options: ConditionalManagerOptions;
  private readonly scope: HTMLSelectElement;
  private readonly list: HTMLDivElement;
  private shown: ManagerRow[] = [];
  /** The place of the first rule shown: the list shows MAX_MANAGER_ROWS at a time. */
  private from = 0;

  constructor(options: ConditionalManagerOptions) {
    this.options = options;
    const strings = options.strings;
    const dialog = new Dialog({ parent: options.parent, title: strings.managerTitle, name: 'conditional-rules', onClose: options.onClose, coarse: options.coarse });
    this.dialog = dialog;
    dialog.panel.style.width = '460px';
    const scope = dialog.select([
      ['selection', strings.scopeSelection],
      ['sheet', strings.scopeSheet],
    ]);
    scope.value = options.scope;
    dialog.addField(dialog.body, strings.showRulesFor, scope, 'scope');
    scope.addEventListener('change', () => {
      this.from = 0;
      this.render();
    });
    this.scope = scope;
    const list = document.createElement('div');
    list.className = 'bym-dialog-rules';
    list.setAttribute('role', 'list');
    dialog.body.appendChild(list);
    this.list = list;
    dialog.addAction(strings.newRule, () => this.leave(() => options.onNew(this.scopeValue())), 'start').setAttribute('data-bym-action', 'new');
    dialog.addAction(strings.close, () => this.submit(), 'primary').setAttribute('data-bym-action', 'close');
    this.render();
    dialog.focusFirst();
  }

  /** Applies what is being typed (the cells of a rule) and closes the list; what cannot be applied keeps it open. */
  submit(): void {
    const root = this.dialog.panel.getRootNode() as Document | ShadowRoot;
    const active = root.activeElement as HTMLInputElement | null;
    if (active && active.getAttribute('data-bym-field') === 'ranges' && this.list.contains(active) && !this.applyRanges(active)) return;
    this.dialog.close();
  }

  /** Whether cells were typed for a rule of the list and not applied yet (they are applied when their field is left). */
  get pending(): boolean {
    if (!this.dialog.isOpen) return false;
    const fields = this.list.querySelectorAll<HTMLInputElement>('input[data-bym-field="ranges"]');
    for (let i = 0; i < fields.length; i++) {
      const index = Number(fields[i].getAttribute('data-bym-index'));
      const row = this.shown.filter((item) => item.index === index)[0];
      if (row && !fields[i].readOnly && fields[i].value !== row.ranges) return true;
    }
    return false;
  }

  private scopeValue(): 'selection' | 'sheet' {
    return this.scope.value === 'sheet' ? 'sheet' : 'selection';
  }

  /** Closes the list for another dialog to take its place. */
  private leave(next: () => void): void {
    this.dialog.close();
    next();
  }

  /** Runs a change; what it refuses is said, and the list is brought up to date either way. Returns whether it went through. */
  private attempt(run: () => void, field?: HTMLElement): boolean {
    this.dialog.showError('');
    try {
      run();
    } catch (problem) {
      this.dialog.showError(problem instanceof FieldError ? problem.message : this.options.explain(problem), field);
      return false;
    }
    return true;
  }

  private applyRanges(input: HTMLInputElement): boolean {
    const index = Number(input.getAttribute('data-bym-index'));
    const row = this.shown.filter((item) => item.index === index)[0];
    if (!row || input.readOnly || input.value === row.ranges) return true;
    const done = this.attempt(() => this.options.onRanges(index, input.value), input);
    if (done) this.render({ index, action: 'ranges' });
    return done;
  }

  /** Lists the rules again; `focus` names the control that keeps the focus, by the rule it belongs to. */
  private render(focus?: { index: number; action: string }): void {
    const strings = this.options.strings;
    let found = this.options.rows(this.scopeValue(), this.from);
    let page = managerPage(found.total, this.from);
    if (page.from !== this.from) {
      // The rules the list began at are gone (deleted, or no longer of the selection): the last ones there are.
      this.from = page.from;
      found = this.options.rows(this.scopeValue(), this.from);
      page = managerPage(found.total, this.from);
    }
    const rows = found.rows.length > MAX_MANAGER_ROWS ? found.rows.slice(0, MAX_MANAGER_ROWS) : found.rows;
    this.shown = rows;
    const list = this.list;
    while (list.firstChild) list.removeChild(list.firstChild);
    if (rows.length === 0) {
      const none = document.createElement('div');
      none.className = 'bym-dialog-hint';
      none.textContent = strings.noRules;
      list.appendChild(none);
    }
    let target: HTMLElement | undefined;
    rows.forEach((row, at) => {
      const item = document.createElement('div');
      item.className = 'bym-dialog-rule';
      item.setAttribute('role', 'listitem');
      item.setAttribute('data-bym-rule', String(at));
      item.setAttribute('aria-label', strings.ruleRow(at + 1, row.description));
      const text = document.createElement('div');
      text.className = 'bym-dialog-rule-text';
      // The words of a rule may come from the file: they are text, never markup.
      text.textContent = row.description;
      item.appendChild(text);
      const line = document.createElement('div');
      line.className = 'bym-dialog-rule-line';
      const ranges = this.dialog.input();
      ranges.value = row.ranges;
      ranges.setAttribute('data-bym-field', 'ranges');
      ranges.setAttribute('data-bym-index', String(row.index));
      ranges.setAttribute('aria-label', `${strings.appliesTo}: ${row.description}`);
      ranges.title = row.adjustable ? strings.appliesTo : row.reason || strings.appliesTo;
      ranges.readOnly = !row.adjustable;
      ranges.addEventListener('keydown', (event) => {
        if (event.key !== 'Enter' || event.isComposing) return;
        // Enter applies the cells of this rule and leaves the list open.
        event.preventDefault();
        event.stopPropagation();
        this.applyRanges(ranges);
      });
      ranges.addEventListener('change', () => {
        if (this.dialog.isOpen) this.applyRanges(ranges);
      });
      line.appendChild(ranges);
      item.appendChild(line);
      const tools = document.createElement('div');
      tools.className = 'bym-dialog-rule-tools';
      const stop = this.dialog.addCheck(tools, strings.stopIfTrue, 'stopIfTrue');
      stop.checked = row.stopIfTrue;
      stop.disabled = !row.adjustable;
      stop.addEventListener('change', () => {
        this.attempt(() => this.options.onStop(row.index, stop.checked));
        this.render({ index: row.index, action: 'stopIfTrue' });
      });
      const button = (name: string, label: string, enabled: boolean, run: () => void, symbol?: string) => {
        const el = document.createElement('button');
        el.type = 'button';
        el.className = symbol ? 'bym-dialog-button bym-dialog-small' : 'bym-dialog-button';
        el.textContent = symbol || label;
        el.setAttribute('data-bym-action', name);
        el.setAttribute('aria-label', `${label}: ${row.description}`);
        el.title = enabled ? label : row.reason || label;
        el.disabled = !enabled;
        el.addEventListener('click', run);
        tools.appendChild(el);
        if (focus && focus.index === row.index && focus.action === name) target = el;
        return el;
      };
      // The first and the last rule shown have their neighbours in the lists before and after.
      const before = at > 0 ? rows[at - 1] : found.before;
      const after = at < rows.length - 1 ? rows[at + 1] : found.after;
      button('up', strings.moveUp, row.movable && !!before && before.movable, () => {
        if (before && this.attempt(() => this.options.onMove(row.index, before.index))) this.render({ index: before.index, action: 'up' });
      }, '\u2191');
      button('down', strings.moveDown, row.movable && !!after && after.movable, () => {
        if (after && this.attempt(() => this.options.onMove(row.index, after.index))) this.render({ index: after.index, action: 'down' });
      }, '\u2193');
      button('edit', strings.edit, row.editable, () => this.leave(() => this.options.onEdit(row.index, this.scopeValue())));
      button('delete', strings.deleteRule, row.removable, () => {
        if (this.attempt(() => this.options.onDelete(row.index))) this.render({ index: -1, action: '' });
      });
      item.appendChild(tools);
      list.appendChild(item);
      if (focus && focus.index === row.index && focus.action === 'ranges') target = ranges;
      if (focus && focus.index === row.index && focus.action === 'stopIfTrue') target = stop;
    });
    if (found.total > rows.length) {
      // More rules than one list holds: which of them these are, and the way to the others.
      const more = document.createElement('div');
      more.className = 'bym-dialog-hint';
      more.setAttribute('data-bym-field', 'more');
      more.textContent = strings.rulesShown(this.from + 1, this.from + rows.length, found.total);
      list.appendChild(more);
      const turn = document.createElement('div');
      turn.className = 'bym-dialog-rule-tools';
      const pageButton = (name: string, label: string, to: number | undefined) => {
        const el = document.createElement('button');
        el.type = 'button';
        el.className = 'bym-dialog-button';
        el.textContent = label;
        el.setAttribute('data-bym-action', name);
        el.disabled = to === undefined;
        el.addEventListener('click', () => {
          if (to === undefined) return;
          // Cells typed for a rule of this list are applied before it is left; what cannot be applied keeps it.
          const fields = list.querySelectorAll<HTMLInputElement>('input[data-bym-field="ranges"]');
          for (let i = 0; i < fields.length; i++) if (!this.applyRanges(fields[i])) return;
          this.dialog.showError('');
          this.from = to;
          this.render({ index: -1, action: name });
        });
        turn.appendChild(el);
        if (focus && focus.action === name) target = el;
      };
      pageButton('previous-page', strings.previousRules(MAX_MANAGER_ROWS), page.previous);
      pageButton('next-page', strings.nextRules(Math.min(MAX_MANAGER_ROWS, found.total - page.to)), page.next);
      list.appendChild(turn);
      this.list.scrollTop = 0;
    }
    if (!focus) return;
    // The control that was used keeps the focus where it can; else the dialog still has it.
    if (target && !(target as HTMLButtonElement).disabled) target.focus({ preventScroll: true });
    else {
      const other = target && target.parentElement ? (target.parentElement.querySelector('button:not(:disabled)') as HTMLElement | null) : null;
      if (other) other.focus({ preventScroll: true });
      else this.dialog.focusFirst();
    }
  }
}

/** The cosmetic rules of the dialogs; what their layout depends on is set inline. */
export const DIALOG_STYLE =
  '.bym-dialog-backdrop{background:rgba(32,33,36,.32);font:13px/1.3 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;letter-spacing:normal;text-transform:none;text-align:left;direction:ltr}' +
  '.bym-dialog{background:#fff;color:#202124;border:1px solid #dadce0;border-radius:8px;box-shadow:0 8px 28px rgba(0,0,0,.28)}' +
  '.bym-dialog-title{padding:12px 16px 8px;font-size:15px;font-weight:600}' +
  '.bym-dialog-body{padding:0 16px 4px}' +
  '.bym-dialog-section{display:block;min-width:0;margin:0 0 10px;padding:6px 10px 8px;border:1px solid #dadce0;border-radius:6px}' +
  '.bym-dialog-legend{float:none;width:auto;margin:0;padding:0 4px;font-size:12px;font-weight:600;color:#5f6368}' +
  '.bym-dialog-row{display:flex;align-items:center;flex-wrap:wrap;margin:6px 0}' +
  '.bym-dialog-label{display:block;flex:0 0 96px;box-sizing:border-box;margin:0;padding-right:8px;font:inherit;color:inherit;overflow-wrap:break-word}' +
  '.bym-dialog-input{-webkit-appearance:none;appearance:none;display:block;flex:1 1 140px;box-sizing:border-box;float:none;width:auto;min-width:0;min-height:0;max-width:none;max-height:none;' +
  'height:30px;margin:0;padding:0 8px;border:1px solid #9aa0a6;border-radius:4px;background:#fff;color:#202124;font:inherit;line-height:1.3;letter-spacing:normal;' +
  'text-transform:none;text-indent:0;text-shadow:none;box-shadow:none;outline:none;opacity:1}' +
  '.bym-dialog-text{height:auto;padding:5px 8px;resize:vertical;white-space:pre-wrap;overflow-wrap:break-word}' +
  // The arrow of a list is drawn with two gradients: no image, and it takes the text colour.
  '.bym-dialog-select{padding-right:24px;cursor:pointer;background-image:linear-gradient(45deg,transparent 50%,currentColor 50%),linear-gradient(135deg,currentColor 50%,transparent 50%);' +
  'background-position:right 12px top 50%,right 7px top 50%;background-size:5px 5px,5px 5px;background-repeat:no-repeat}' +
  '.bym-dialog-input:focus{border-color:#1a7f45;box-shadow:0 0 0 1px #1a7f45}' +
  '.bym-dialog-input[aria-invalid="true"]{border-color:#d93025;box-shadow:0 0 0 1px #d93025}' +
  '.bym-dialog-check{display:flex;align-items:center;margin:6px 0;font:inherit;color:inherit;cursor:pointer}' +
  '.bym-dialog-box{flex:none;box-sizing:border-box;width:16px;height:16px;min-width:0;min-height:0;margin:0 8px 0 0;padding:0;opacity:1;accent-color:#1a7f45}' +
  '.bym-dialog-hint,.bym-dialog-note{margin:2px 0 6px;color:#5f6368;font-size:12px;overflow-wrap:break-word}' +
  '.bym-dialog-note{padding:6px 8px;background:#fef7e0;border-radius:4px;color:#202124}' +
  '.bym-dialog-alert{padding:0 16px;color:#b3261e;font-size:12px;overflow-wrap:break-word}' +
  '.bym-dialog-alert:empty{display:none}' +
  '.bym-dialog-actions{padding:8px 12px 12px}' +
  '.bym-dialog-button{-webkit-appearance:none;appearance:none;display:block;flex:none;box-sizing:border-box;float:none;width:auto;height:32px;min-width:64px;min-height:0;max-width:none;max-height:none;' +
  'margin:4px;padding:0 14px;border:1px solid #dadce0;border-radius:6px;background:#fff;color:#202124;font:inherit;line-height:1.2;letter-spacing:normal;text-transform:none;' +
  'text-shadow:none;opacity:1;cursor:pointer}' +
  '.bym-dialog-primary{background:#1a7f45;border-color:#1a7f45;color:#fff}' +
  '.bym-dialog-button:focus{outline:2px solid #1a7f45;outline-offset:1px}' +
  '.bym-dialog-button:disabled{opacity:.45;cursor:default}' +
  '.bym-dialog-color{-webkit-appearance:none;appearance:none;display:block;flex:none;box-sizing:border-box;float:none;width:48px;height:30px;min-width:0;min-height:0;margin:0;padding:2px;' +
  'border:1px solid #9aa0a6;border-radius:4px;background:#fff;cursor:pointer;opacity:1}' +
  '.bym-dialog-color:focus{outline:2px solid #1a7f45;outline-offset:1px}' +
  '.bym-dialog-rules{margin:6px 0 4px}' +
  '.bym-dialog-rule{margin:0 0 8px;padding:6px 8px;border:1px solid #dadce0;border-radius:6px}' +
  '.bym-dialog-rule-text{font-weight:600;overflow-wrap:break-word;word-break:break-word}' +
  '.bym-dialog-rule-line{display:flex;margin:6px 0 2px}' +
  '.bym-dialog-rule-tools{display:flex;flex-wrap:wrap;align-items:center}' +
  '.bym-dialog-rule-tools .bym-dialog-check{margin:0 auto 0 0;padding-right:8px}' +
  '.bym-dialog-rule-tools .bym-dialog-button{min-width:0;height:28px;margin:2px;padding:0 10px}' +
  '.bym-dialog-small{width:30px}' +
  '.bym-dialog-input[readonly]{background:#f1f3f4;color:#5f6368}' +
  // A finger needs more to hit than a mouse pointer.
  '@media (pointer:coarse){.bym-dialog-color{height:40px;width:56px}.bym-dialog-rule-tools .bym-dialog-button{height:40px}.bym-dialog-small{width:40px}}' +
  '@media (pointer:coarse){.bym-dialog-input{height:40px}.bym-dialog-text{height:auto}.bym-dialog-button{height:40px}.bym-dialog-box{width:22px;height:22px}.bym-dialog-check{min-height:40px}}';
