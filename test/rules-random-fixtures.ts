// A seeded mix of everything that touches rules (data validations, conditional formats):
// rules set, changed and cleared, rows and columns inserted and deleted, sheets renamed,
// deleted, copied, moved and added, cells written, undo and redo. Used by the random test and
// by the script that hands the files it makes to another reader.

import type { CellRange, DataValidation, Workbook } from '../src/index';
import { rangeName } from '../src/index';
import { DataValidations } from '../src/format/validation';
import type { History, Op } from '../src/ops/ops';
import { CalcEngine } from '../src/calc/engine';
import { ConditionalFormats } from '../src/format/conditional';
import {
  addConditionalRule,
  clearConditionalRules,
  conditionalRules,
  dataValidationNow,
  moveConditionalRule,
  removeConditionalRule,
  ruleSpace,
  setValidation,
  updateConditionalRule,
  type ConditionalSettings,
  type ValidationSettings,
} from '../src/ops/rules';
import { LOCALE_EN, LOCALE_TR } from '../src/format/locale';
import { resolveColor } from '../src/model/sheet/color';
import { validationsIn } from '../src/ops/rules';
import { conditionalFields, conditionalFromFields, conditionalPatch, rangeListText, validationChange, validationFields, validationFromFields } from '../src/ui/rules-dialog';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';
import { BLOCK_A, BLOCK_B, BLOCK_C, BLOCK_D, EXT_BAR, EXT_ONLY, REVIEW_DXFS, STYLES, UNMODELLED_FORMATS, UNMODELLED_VALIDATION, cf, conditionalContent, reviewRules } from './conditional-fixtures';
import { THEME } from './sheet-drawing-fixtures';
import { copySheet, deleteSheet, insertSheet, moveSheet, renameSheet } from '../src/ops/sheets';
import { buildXlsx, worksheetXml } from './fixtures';

/** A small generator of numbers in [0, 1) that gives the same run for the same seed. */
export function seeded(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const dv = (attrs: string, inner = '') => `<dataValidation ${attrs}>${inner}</dataValidation>`;

/**
 * Rules of kinds that look at their own cell, their formulas spelled with "$" as other
 * producers write them: the "$" anchors nothing there (see ownCellProblems()).
 */
const OWN_CELL_RULES =
  cf('D3:D7 F1:F3', '<cfRule type="timePeriod" dxfId="2" priority="2" timePeriod="last7Days"><formula>AND(TODAY()-FLOOR($D3,1)&lt;=6,FLOOR($D3,1)&lt;=TODAY())</formula></cfRule>') +
  cf('E1:E4', '<cfRule type="containsText" dxfId="0" priority="3" operator="containsText" text="x"><formula>NOT(ISERROR(SEARCH("x",$E$1)))</formula></cfRule>') +
  cf('G2:G5 C8:C9', '<cfRule type="containsBlanks" dxfId="1" priority="4"><formula>LEN(TRIM(G$2))=0</formula></cfRule>');

const CELL = '(\\$?[A-Z]{1,3}\\$?[0-9]{1,7})';
const quotedText = (text: string) => `"${text.replace(/"/g, '""')}"`;
/** The formula Excel writes for each kind of rule that looks at its own cell, with `c` where the cell stands. Written out here, apart from the engine. */
const OWN_CELL_FORMULAS: Record<string, (c: string, text: string, period: string) => string | undefined> = {
  containsText: (c, text) => `NOT(ISERROR(SEARCH(${quotedText(text)},${c})))`,
  notContainsText: (c, text) => `ISERROR(SEARCH(${quotedText(text)},${c}))`,
  beginsWith: (c, text) => `LEFT(${c},LEN(${quotedText(text)}))=${quotedText(text)}`,
  endsWith: (c, text) => `RIGHT(${c},LEN(${quotedText(text)}))=${quotedText(text)}`,
  containsBlanks: (c) => `LEN(TRIM(${c}))=0`,
  notContainsBlanks: (c) => `LEN(TRIM(${c}))>0`,
  containsErrors: (c) => `ISERROR(${c})`,
  notContainsErrors: (c) => `NOT(ISERROR(${c}))`,
  timePeriod: (c, _text, period) =>
    ({
      today: `FLOOR(${c},1)=TODAY()`,
      yesterday: `FLOOR(${c},1)=TODAY()-1`,
      tomorrow: `FLOOR(${c},1)=TODAY()+1`,
      last7Days: `AND(TODAY()-FLOOR(${c},1)<=6,FLOOR(${c},1)<=TODAY())`,
      lastWeek: `AND(TODAY()-ROUNDDOWN(${c},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${c},0)<(WEEKDAY(TODAY())+7))`,
      thisWeek: `AND(TODAY()-ROUNDDOWN(${c},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${c},0)-TODAY()<=7-WEEKDAY(TODAY()))`,
      nextWeek: `AND(ROUNDDOWN(${c},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${c},0)-TODAY()<(15-WEEKDAY(TODAY())))`,
      lastMonth: `AND(MONTH(${c})=MONTH(EDATE(TODAY(),0-1)),YEAR(${c})=YEAR(EDATE(TODAY(),0-1)))`,
      thisMonth: `AND(MONTH(${c})=MONTH(TODAY()),YEAR(${c})=YEAR(TODAY()))`,
      nextMonth: `AND(MONTH(${c})=MONTH(EDATE(TODAY(),0+1)),YEAR(${c})=YEAR(EDATE(TODAY(),0+1)))`,
    })[period],
};

const xmlText = (text: string) => text.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/**
 * What a sheet part says wrong about the rules that look at their own cell: for each <cfRule>
 * of such a kind, what its attributes say (the kind, the text, the period), the cell its
 * formula names and the first cell of its `sqref` are one statement. A formula that is the one
 * of the kind names that first cell, with whatever "$"; and none of them names a cell that is
 * gone. A formula that is not the one of the kind (written by hand) is not looked at.
 */
export function ownCellProblems(xml: string): string[] {
  const out: string[] = [];
  const attr = (tag: string, name: string) => {
    const found = new RegExp(`\\s${name}=(?:"([^"]*)"|'([^']*)')`).exec(tag);
    return found ? xmlText(found[1] !== undefined ? found[1] : found[2]) : undefined;
  };
  for (const block of xml.match(/<conditionalFormatting\b[^>]*>[\s\S]*?<\/conditionalFormatting>/g) || []) {
    const sqref = attr(block.slice(0, block.indexOf('>') + 1), 'sqref') || '';
    const first = sqref.split(' ')[0].split(':')[0];
    for (const rule of block.match(/<cfRule\b[^>]*>[\s\S]*?<\/cfRule>/g) || []) {
      const tag = rule.slice(0, rule.indexOf('>') + 1);
      const make = OWN_CELL_FORMULAS[attr(tag, 'type') || ''];
      const formulas = (rule.match(/<formula>[\s\S]*?<\/formula>/g) || []).map((text) => xmlText(text.slice(9, -10)));
      if (!make || formulas.length !== 1) continue;
      if (formulas[0].indexOf('#REF!') >= 0) out.push(`${sqref}: ${formulas[0]} names a cell that is gone`);
      // The formula of the kind with a place for the cell: every such place is to name one and the same cell.
      const pattern = make('\u0000', attr(tag, 'text') || '', attr(tag, 'timePeriod') || '');
      if (pattern === undefined) continue;
      const found = new RegExp(`^${pattern.split('\u0000').map((piece) => piece.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join(CELL)}$`).exec(formulas[0]);
      if (!found) continue;
      const cells = found.slice(1);
      if (cells.some((cell) => cell !== cells[0]) || cells[0].replace(/\$/g, '') !== first) out.push(`${sqref}: ${formulas[0]} does not name ${first}`);
    }
  }
  return out;
}

/** Three sheets; the first has rules of every kind of place: several ranges, another sheet, the extension list. */
export function randomBook(options: { conditional?: boolean; extensionOnly?: boolean } = {}): Promise<Uint8Array> {
  const data =
    '<sheetData><row r="1"><c r="A1"><v>5</v></c><c r="B1"><v>1</v></c><c r="H1"><v>1</v></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>7</v></c><c r="H2"><v>2</v></c></row><row r="3"><c r="A3"><v>5</v></c><c r="C3"><v>4</v></c></row></sheetData>' +
    (options.conditional ? BLOCK_A + BLOCK_B + BLOCK_C + BLOCK_D : '') +
    '<dataValidations count="4">' +
    dv('type="whole" showErrorMessage="1" errorTitle="Range" error="1 to 9" sqref="A1:A10 C1:C3"', '<formula1>1</formula1><formula2>9</formula2>') +
    dv('type="list" allowBlank="1" showInputMessage="1" promptTitle="Pick" prompt="One of them" sqref="B1:B5"', '<formula1>$H$1:$H$2</formula1>') +
    dv('type="custom" showErrorMessage="1" errorStyle="warning" sqref="D2:E6"', '<formula1>AND(D2&gt;$H$1,Lists!$A$1&lt;&gt;D1)</formula1>') +
    dv('type="list" sqref="F1:F4"', '<formula1>Lists!$A$1:$A$3</formula1>') +
    '</dataValidations>' +
    '<extLst>' +
    (options.conditional ? `<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings>${EXT_BAR}${options.extensionOnly ? EXT_ONLY : ''}</x14:conditionalFormattings></ext>` : '') +
    '<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
    '<x14:dataValidations xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" count="1">' +
    '<x14:dataValidation type="list" allowBlank="1" showErrorMessage="1"><x14:formula1><xm:f>Lists!$B$1:$B$3</xm:f></x14:formula1><xm:sqref>L1:L3</xm:sqref></x14:dataValidation>' +
    '</x14:dataValidations></ext></extLst>';
  const lists =
    '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>4</v></c></row></sheetData>' +
    (options.conditional ? cf('A1:B3', '<cfRule type="cellIs" dxfId="1" priority="1" operator="equal"><formula>Data!$H$1</formula></cfRule>') + OWN_CELL_RULES : '') +
    '<dataValidations count="1">' +
    dv('type="decimal" operator="greaterThan" sqref="A1:B3"', '<formula1>Data!$H$1</formula1>') +
    '</dataValidations>';
  return buildXlsx({
    sheets: [
      { name: 'Data', xml: worksheetXml(data) },
      { name: 'Lists', xml: worksheetXml(lists) },
      // Rules that say more than a dialog shows, elements the reader leaves out of the model, a validation with attributes of its own.
      {
        name: 'Other',
        xml: worksheetXml(
          '<sheetData/>' +
            (options.conditional ? UNMODELLED_FORMATS + reviewRules(3) : '') +
            `<dataValidations count="3" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">${UNMODELLED_VALIDATION}` +
            dv('type="list" imeMode="off" xr:uid="{CCCC}" allowBlank="1" sqref="K1:K5"', '<formula1>"a, b,c"</formula1>') +
            dv('type="whole" sqref="K7:K8"', '<formula1>SUM( $A$1:$A$3 )</formula1><formula2>99</formula2>') +
            '</dataValidations>',
        ),
      },
    ],
    ...(options.conditional ? { styles: STYLES.replace(/<dxfs count="3">(.*)<\/dxfs>/, `<dxfs count="7">$1${REVIEW_DXFS}</dxfs>`), theme: THEME } : {}),
  });
}

export interface RandomOptions {
  /** Only rules are changed: no cell, row, column or sheet. */
  rulesOnly?: boolean;
  /** Conditional formats are changed as well as data validations. */
  conditional?: boolean;
}

const COLORS = ['#FFC7CE', '#9C0006', '#FFEB9C', '#C6EFCE', '#006100', '#FF0000'];

/** What one step did, for the message of a failed comparison. */
export type Step = {
  what: string;
  ops?: Op[];
  undo?: boolean;
  redo?: boolean;
  /** The same command for another workbook in the same state. */
  build?: (wb: Workbook) => Op[];
  /** For a step that sets or clears the rules of cells: the sheet and the cells, and whether a validation is there afterwards. */
  cells?: { kind: 'validation' | 'conditionalClear'; sheet: number; ranges: CellRange[] | undefined; set?: boolean };
};

const TITLES = ['', 'Title', 'Başlık & <b>'];
const TEXTS = ['', 'Plain', 'Two\nlines', 'Quote " and \' and _x000a_ kept'];

/** One random step: the operations to apply (through the history), or an undo or a redo. Undefined when the step came to nothing. */
export function randomStep(wb: Workbook, history: History, random: () => number, options: RandomOptions = {}): Step | undefined {
  const pick = <T>(list: readonly T[]): T => list[Math.floor(random() * list.length)];
  const int = (limit: number) => Math.floor(random() * limit);
  const worksheets = wb.sheets.map((sheet, index) => ({ sheet, index })).filter((entry) => entry.sheet.kind === 'worksheet');
  const target = pick(worksheets);
  const others = wb.sheets.filter((sheet) => sheet !== target.sheet).map((sheet) => sheet.name);
  const quoted = (name: string) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`);
  const range = (): CellRange => {
    const top = int(14);
    const left = int(13);
    return { top, left, bottom: top + int(4), right: left + int(3) };
  };
  const refTo = (): string => {
    // Now and then a cell at the far edge of the sheet: moved with its rule, such a reference goes round the edge.
    const cell = random() < 0.08 ? pick(['A1048576', 'XFD1', 'B$1048575', '$XFC3']) : `${pick(['', '$'])}${'ABCDEFGH'.charAt(int(8))}${pick(['', '$'])}${1 + int(12)}`;
    return others.length > 0 && random() < 0.4 ? `${quoted(pick(others))}!${cell}` : cell;
  };
  const roll = random();
  const structural = !options.rulesOnly;
  // The command is built for this workbook here, so that what it refuses is no step; `build` makes it again for another.
  const made = (what: string, build: (book: Workbook) => Op[]): Step => ({ what, ops: build(wb), build });
  try {
    if (roll < 0.12) return history.canUndo ? { what: 'undo', undo: true } : undefined;
    if (roll < 0.17) return history.canRedo ? { what: 'redo', redo: true } : undefined;
    if (options.conditional && random() < 0.6 && (roll < 0.52 || !structural)) {
      const listed = conditionalRules(wb, target.index);
      const at = random() < 0.2 ? [range(), range()] : range();
      const format = random() < 0.85 ? { ...(random() < 0.6 ? { fill: pick(COLORS) } : {}), ...(random() < 0.5 ? { color: pick(COLORS) } : {}), ...(random() < 0.3 ? { bold: random() < 0.7 } : {}), ...(random() < 0.2 ? { italic: true } : {}), ...(random() < 0.15 ? { underline: random() < 0.8 } : {}), ...(random() < 0.1 ? { strike: true } : {}) } : undefined;
      const settings = (): ConditionalSettings => {
        const type = pick(['cellIs', 'cellIs', 'containsText', 'notContainsText', 'beginsWith', 'endsWith', 'duplicateValues', 'uniqueValues', 'top10', 'aboveAverage', 'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors', 'expression', 'expression', 'colorScale', 'dataBar'] as const);
        const out: ConditionalSettings = { type, stopIfTrue: random() < 0.2, format };
        if (type === 'cellIs') {
          out.operator = pick(['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'] as const);
          out.formulas = [random() < 0.5 ? String(int(9)) : refTo(), random() < 0.5 ? String(9 + int(9)) : refTo()];
        } else if (type === 'expression') out.formulas = [`AND(${refTo()}>${int(9)},${refTo()}<>"x")`];
        else if (type === 'top10') Object.assign(out, { rank: 1 + int(20), percent: random() < 0.3, bottom: random() < 0.3 });
        else if (type === 'aboveAverage') Object.assign(out, { aboveAverage: random() < 0.5, equalAverage: random() < 0.3 });
        else if (type === 'colorScale') out.colors = random() < 0.5 ? [pick(COLORS), pick(COLORS)] : [pick(COLORS), pick(COLORS), pick(COLORS)];
        else if (type === 'dataBar') out.colors = [pick(COLORS)];
        else if (type !== 'duplicateValues' && type !== 'uniqueValues') out.text = pick(['a', 'x"y', 'Şube & <b>']);
        return out;
      };
      const kind = random();
      if (kind < 0.4 || listed.length === 0) {
        const rule = settings();
        return made(`conditional add ${rule.type}`, (book) => addConditionalRule(book, target.index, at, rule));
      }
      const index = pick(listed).index;
      if (kind < 0.55) {
        const patch = random() < 0.4 ? { ranges: at } : random() < 0.5 ? { stopIfTrue: random() < 0.5 } : { rule: settings() };
        return made('conditional update', (book) => updateConditionalRule(book, target.index, index, patch));
      }
      if (kind < 0.7) return made('conditional remove', (book) => removeConditionalRule(book, target.index, index));
      if (kind < 0.85) {
        const to = int(listed.length);
        return made('conditional move', (book) => moveConditionalRule(book, target.index, index, to));
      }
      const whole = random() < 0.15;
      return { ...made('conditional clear', (book) => clearConditionalRules(book, target.index, whole ? undefined : at)), cells: { kind: 'conditionalClear' as const, sheet: target.index, ranges: whole ? undefined : Array.isArray(at) ? at : [at] } };
    }
    if (roll < 0.52 || !structural) {
      const type = pick(['whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom', 'none', 'clear'] as const);
      const at = random() < 0.25 ? [range(), range()] : range();
      const cellsOf = (set: boolean) => ({ kind: 'validation' as const, sheet: target.index, ranges: Array.isArray(at) ? at : [at], set });
      if (type === 'clear') return { ...made('clear validation', (book) => setValidation(book, target.index, at, undefined)), cells: cellsOf(false) };
      const operator = pick(['between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'] as const);
      const settings: ValidationSettings = {
        type,
        operator,
        allowBlank: random() < 0.5,
        hideDropDown: random() < 0.3,
        showInputMessage: random() < 0.6,
        showErrorMessage: random() < 0.7,
        errorStyle: pick(['stop', 'warning', 'information'] as const),
        promptTitle: pick(TITLES),
        prompt: pick(TEXTS),
        errorTitle: pick(TITLES),
        error: pick(TEXTS),
      };
      if (type === 'list') settings.formula1 = random() < 0.5 ? '"a,b ""c"",d"' : `${refTo()}:${refTo().replace(/^.*!/, '')}`;
      else if (type === 'custom') settings.formula1 = `AND(${refTo()}>${int(9)},${refTo()}<>"x")`;
      else if (type !== 'none') {
        settings.formula1 = random() < 0.5 ? String(int(50)) : refTo();
        settings.formula2 = random() < 0.5 ? String(50 + int(50)) : refTo();
      }
      // "Any value" with a message to show is a rule; without one it clears.
      const stays = type !== 'none' || (settings.showInputMessage === true && (settings.promptTitle !== '' || settings.prompt !== ''));
      return { ...made(`validation ${type}`, (book) => setValidation(book, target.index, at, settings)), cells: cellsOf(stays) };
    }
    if (roll < 0.74) {
      const axis = pick(['row', 'col'] as const);
      const kind = pick(['insert', 'delete'] as const);
      const op: Op = { type: kind, sheet: target.index, axis, at: int(12), count: 1 + int(3) };
      return made(`${kind} ${axis}`, () => [op]);
    }
    if (roll < 0.82) {
      const op: Op = { type: 'setCells', sheet: target.index, cells: [{ row: int(12), col: int(8), cell: { kind: 'number', value: int(100), style: 0 } }] };
      return made('cell', () => [op]);
    }
    if (roll < 0.87) {
      const name = pick(['Data', 'Lists', 'Other', 'My Sheet', "It's", 'Şube', 'S1', 'S2']);
      return made('rename', (book) => renameSheet(book, target.index, name));
    }
    if (roll < 0.9) return wb.sheets.length > 2 ? made('delete sheet', (book) => deleteSheet(book, target.index)) : undefined;
    if (roll < 0.94) return wb.sheets.length < 6 ? made('copy sheet', (book) => copySheet(book, target.index)) : undefined;
    if (roll < 0.97) {
      const at = int(wb.sheets.length + 1);
      const name = `New${int(1000)}`;
      return wb.sheets.length < 6 ? made('insert sheet', (book) => insertSheet(book, at, { name })) : undefined;
    }
    const to = int(wb.sheets.length);
    return made('move sheet', (book) => moveSheet(book, target.index, to));
  } catch (error) {
    // What a command refuses (a name that is taken, cells under a rule of the extension list) is no step.
    if (error instanceof RangeError) return undefined;
    throw error;
  }
}

const message = (rule: DataValidation) => [rule.allowBlank, rule.hideDropDown, rule.showInputMessage, rule.promptTitle, rule.prompt, rule.showErrorMessage, rule.errorStyle, rule.errorTitle, rule.error];

/** The data validations of every sheet as they apply now, in order, as plain data. */
export function validationsNow(wb: Workbook): unknown[] {
  return wb.sheets.map((sheet) => {
    const space = ruleSpace(wb, sheet);
    return [
      sheet.name,
      (sheet.dataValidations || [])
        .map((rule) => dataValidationNow(rule, space))
        .filter((rule): rule is DataValidation => !!rule)
        .map((rule) => [rule.ranges.map(rangeName).join(' '), rule.type, rule.operator, rule.formula1, rule.formula2, rule.extended === true, ...message(rule)]),
    ];
  });
}

/** What the evaluation says about a block of cells of every sheet: the rule each cell has, as it is looked up. */
export function validationsSeen(wb: Workbook): unknown[] {
  return wb.sheets.map((sheet) => {
    const rules = new DataValidations(wb, sheet);
    const cells: string[] = [];
    for (let row = 0; row < 22; row++) {
      for (let col = 0; col < 18; col++) {
        const rule = rules.ruleAt(row, col);
        if (rule) cells.push(`${row},${col}:${rule.type}|${rule.formula1 ?? ''}|${rule.formula2 ?? ''}|${rule.ranges.map(rangeName).join(' ')}`);
      }
    }
    return cells;
  });
}

/** What the conditional formats give a block of cells of every sheet, with formulas calculated. */
export function conditionalSeen(wb: Workbook): unknown[] {
  const engine = new CalcEngine(wb);
  return wb.sheets.map((sheet) => {
    const formats = new ConditionalFormats(wb, sheet, { engine: () => engine, budget: 0 });
    const cells: string[] = [];
    for (let row = 0; row < 22; row++) {
      for (let col = 0; col < 18; col++) {
        const found = formats.resolve(row, col);
        if (found) cells.push(`${row},${col}:${JSON.stringify(found)}`);
      }
    }
    return cells;
  });
}

/** Everything the rules of a workbook say, as they apply now: plain data to compare. */
export const rulesNow = (wb: Workbook): unknown => [validationsNow(wb), conditionalContent(wb)];
/** Everything the rules of a workbook do to a block of cells of every sheet. */
export const rulesSeen = (wb: Workbook): unknown => [validationsSeen(wb), conditionalSeen(wb)];
/** What a file built from the model keeps of the rules: the validations as they are looked up, the conditional rules the main part of a sheet states. */
export const rulesBuilt = (wb: Workbook): unknown => [validationsSeen(wb), conditionalContent(wb, true)];

/**
 * Every rule of the workbook, shown in a dialog and typed back without a change, is no step:
 * neither by what the dialog makes of fields nobody touched, nor by the rule its fields stand
 * for when it is set again on the cells it has. Returns what is wrong, or undefined.
 */
export function restatedProblem(wb: Workbook): string | undefined {
  for (const [locale, strings] of [[LOCALE_EN, STRINGS_EN], [LOCALE_TR, STRINGS_TR]] as const) {
    for (let s = 0; s < wb.sheets.length; s++) {
      if (wb.sheets[s].kind !== 'worksheet') continue;
      for (const info of conditionalRules(wb, s)) {
        if (!info.editable) continue;
        const where = `${wb.sheets[s].name}, rule ${info.index} (${info.rule.type})`;
        // Colours as the view shows them: those of the theme through the style sheet.
        const fields = conditionalFields(info, locale, false, undefined, (color) => resolveColor(color, wb.styles));
        const ranges = rangeListText(info.ranges, locale);
        if (Object.keys(conditionalPatch({ ...fields }, ranges, { fields, ranges }, info, locale, strings.rules)).length > 0) return `${where}: the dialog changes something nobody touched`;
        // The whole of what the fields say, stated again. (A text rule with a formula of its own is shown as that formula:
        // stated in full it would be a formula rule, which is why the dialog only states what was touched.)
        if (fields.type !== info.rule.type) continue;
        const ops = updateConditionalRule(wb, s, info.index, { ranges: info.ranges, rule: conditionalFromFields(fields, locale, strings.rules), stopIfTrue: fields.stopIfTrue });
        if (ops.length > 0) return `${where}: stated again from its fields it is another rule`;
      }
      const space = ruleSpace(wb, wb.sheets[s]);
      for (const stated of wb.sheets[s].dataValidations || []) {
        const rule = dataValidationNow(stated, space);
        if (!rule || rule.extended || validationsIn(wb, s, rule.ranges).length !== 1) continue;
        const where = `${wb.sheets[s].name}, validation of ${rule.ranges.map(rangeName).join(' ')}`;
        const fields = validationFields(rule, locale, false);
        if (validationChange({ ...fields }, { fields, rule }, locale, strings.rules) !== undefined) return `${where}: the dialog changes something nobody touched`;
        // Typed back in full: the same rule, but for the spelling of a value (which the dialog keeps when it was not touched).
        const typed = validationFromFields(fields, locale, strings.rules);
        const kept = { ...typed, formula1: typed.formula1 === undefined ? undefined : rule.formula1, formula2: typed.formula2 === undefined ? undefined : rule.formula2 };
        if (setValidation(wb, s, rule.ranges, kept).length > 0) return `${where}: stated again from its fields it is another rule`;
      }
    }
  }
  return undefined;
}

/** For a block of cells of a sheet: whether a validation applies and what it says there (the verdict on a number, the list offered). */
export function validationBehaviour(wb: Workbook, sheetIndex: number): Map<string, string> {
  const engine = new CalcEngine(wb);
  const rules = new DataValidations(wb, wb.sheets[sheetIndex], { engine: () => engine });
  const out = new Map<string, string>();
  for (let row = 0; row < 24; row++) {
    for (let col = 0; col < 18; col++) {
      if (!rules.ruleAt(row, col)) continue;
      const list = rules.list(row, col);
      out.set(`${row},${col}`, `${rules.check(row, col, { kind: 'number', value: 5, style: 0 }, '5').ok}|${rules.check(row, col, { kind: 'string', value: 'a', style: 0 }, 'a').ok}|${list ? list.items.join(',') : '-'}`);
    }
  }
  return out;
}

/** For a block of cells of a sheet: how many conditional rules apply to each. */
export function conditionalCover(wb: Workbook, sheetIndex: number): Map<string, number> {
  const out = new Map<string, number>();
  for (const info of conditionalRules(wb, sheetIndex)) {
    for (const range of info.ranges) {
      for (let row = range.top; row <= Math.min(range.bottom, 23); row++) for (let col = range.left; col <= Math.min(range.right, 17); col++) out.set(`${row},${col}`, (out.get(`${row},${col}`) || 0) + 1);
    }
  }
  return out;
}
