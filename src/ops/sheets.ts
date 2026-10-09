// Operations on the sheets of a workbook as the editor asks for them: a new
// sheet, a sheet deleted, renamed, moved, hidden or shown, copied, and the
// colour of its tab. Like the commands on cells these
// functions only build operations; the caller applies them through a History
// so that each is one undo step. What follows from them (formulas, defined
// names and links that name the sheet) is worked out when they are applied.

import { renameSheetsInFormula } from '../calc/refs';
import { protectionNow } from '../formats/xlsx/build';
import { commentCellNow } from '../render/sheet-drawings';
import { conditionalFormatNow, dataValidationNow, extensionRuleCount, mainPartRule, ruleSpace, ruleWithFormulas } from './rules';
import { MAX_COLS, MAX_ROWS, checkSheetName, createSheet, foldSheetName, type Cell, type Color, type ConditionalFormat, type ConditionalRule, type DataValidation, type DefinedName, type Hyperlink, type Row, type Sheet, type SheetComment, type SheetNameMap, type SheetState, type Workbook } from '../model/sheet/types';
import { copySheetProblem, deleteSheetProblem, insertSheetProblem, moveSheetProblem, renameSheetProblem, sheetStateProblem, tabColorProblem, type Op } from './ops';

export interface InsertSheetOptions {
  /** The name of the new sheet; the next free default name of the language without it. */
  name?: string;
  /** The language default names are in, as a tag ("tr", "tr-TR", "en"): Turkish gives "Sayfa1", every other "Sheet1". */
  locale?: string;
}

/** Highest number a default name is tried with; a workbook never has that many sheets. */
const MAX_DEFAULT_NUMBER = 1000000;

/**
 * The name a new sheet gets when none is given: "Sheet" (Turkish: "Sayfa") with the lowest
 * number from 1 that no sheet of the workbook has, whatever the case of its letters.
 */
export function defaultSheetName(workbook: Workbook, locale?: string): string {
  const base = typeof locale === 'string' && /^tr(?:[-_]|$)/i.test(locale.trim()) ? 'Sayfa' : 'Sheet';
  const taken = new Set<string>();
  for (const sheet of workbook.sheets) taken.add(foldSheetName(sheet.name));
  for (let n = 1; n <= MAX_DEFAULT_NUMBER; n++) {
    const name = base + n;
    if (!taken.has(foldSheetName(name))) return name;
  }
  throw new RangeError('The workbook has too many sheets');
}

/**
 * Why the sheet at `index` cannot be renamed, as a message; undefined when it can. With
 * `name` the name is checked as well (its characters and length, and that no other sheet
 * has it). Cheap: no cell is looked at, so a formula that cannot be read is only found
 * when the operation is applied.
 */
export function canRenameSheet(workbook: Workbook, index: number, name?: string): string | undefined {
  return renameSheetProblem(workbook, index, name);
}

/** Why the sheet at `index` cannot be deleted, as a message; undefined when it can. Cheap, like canRenameSheet(). */
export function canDeleteSheet(workbook: Workbook, index: number): string | undefined {
  return deleteSheetProblem(workbook, index);
}

/** Why no sheet can be inserted at `at` under `name` (or under the default name), as a message; undefined when one can. */
export function canInsertSheet(workbook: Workbook, at: number, name?: string): string | undefined {
  return insertSheetProblem(workbook, at, name === undefined ? defaultSheetName(workbook) : name);
}

/**
 * Renames the sheet at `index`. Formulas, defined names and links of the whole workbook that
 * name the sheet are written with the new name when the operation is applied. Nothing to do
 * (an empty list) when the sheet has that name already. Throws RangeError for a name a sheet
 * cannot have or another sheet has, and for a sheet that cannot be renamed.
 */
export function renameSheet(workbook: Workbook, index: number, name: string): Op[] {
  const problem = renameSheetProblem(workbook, index, name);
  if (problem) throw new RangeError(problem);
  if (workbook.sheets[index].name === name) return [];
  return [{ type: 'renameSheet', sheet: index, name }];
}

/** Deletes the sheet at `index`. Throws RangeError for a sheet that cannot be deleted (the last visible one, say). */
export function deleteSheet(workbook: Workbook, index: number): Op[] {
  const problem = deleteSheetProblem(workbook, index);
  if (problem) throw new RangeError(problem);
  return [{ type: 'deleteSheet', sheet: index }];
}

/**
 * Inserts an empty worksheet so that it is sheet number `at` (0: first; the number of
 * sheets: last). Throws RangeError for an index outside that and for a name that cannot be used.
 */
export function insertSheet(workbook: Workbook, at: number, options: InsertSheetOptions = {}): Op[] {
  const name = options.name !== undefined ? options.name : defaultSheetName(workbook, options.locale);
  const problem = insertSheetProblem(workbook, at, name);
  if (problem) throw new RangeError(problem);
  return [{ type: 'insertSheet', sheet: at, name }];
}

/** Why the sheet at `from` cannot be moved to be sheet number `to`, as a message; undefined when it can. */
export function canMoveSheet(workbook: Workbook, from: number, to: number): string | undefined {
  return moveSheetProblem(workbook, from, to);
}

/**
 * Moves the sheet at `from` so that it is sheet number `to` (0: first; the number of sheets
 * less one: last). Nothing to do when it is there already. Formulas are not changed; defined
 * names that are local to a sheet and the active sheet follow their sheets.
 */
export function moveSheet(workbook: Workbook, from: number, to: number): Op[] {
  const problem = moveSheetProblem(workbook, from, to);
  if (problem) throw new RangeError(problem);
  return from === to ? [] : [{ type: 'moveSheet', sheet: from, to }];
}

/** Why the sheet at `index` cannot be hidden, as a message; undefined when it can (the last visible sheet cannot). */
export function canHideSheet(workbook: Workbook, index: number): string | undefined {
  return sheetStateProblem(workbook, index, 'hidden');
}

/**
 * Shows ("visible") or hides ("hidden") the sheet at `index`; "veryHidden" hides it so that
 * applications do not offer to show it. Nothing to do when the sheet is in that state.
 * Throws RangeError for the last visible sheet.
 */
export function setSheetState(workbook: Workbook, index: number, state: SheetState): Op[] {
  const problem = sheetStateProblem(workbook, index, state);
  if (problem) throw new RangeError(problem);
  return workbook.sheets[index].state === state ? [] : [{ type: 'setSheetState', sheet: index, state }];
}

/** A colour as "#RRGGBB" for the tab of a sheet: the colour of the model for it. Throws RangeError for anything else. */
function tabColorOf(color: string): Color {
  const match = typeof color === 'string' ? /^#([0-9a-fA-F]{6})$/.exec(color) : null;
  if (!match) throw new RangeError('A tab colour is written #RRGGBB');
  return { rgb: `FF${match[1].toUpperCase()}` };
}

/**
 * Sets the colour of the tab of the sheet at `index` ("#RRGGBB"); undefined removes it.
 * Nothing to do when the tab has that colour (or none) already.
 */
export function setTabColor(workbook: Workbook, index: number, color: string | undefined): Op[] {
  const problem = tabColorProblem(workbook, index);
  if (problem) throw new RangeError(problem);
  const wanted = color === undefined || color === null ? undefined : tabColorOf(color);
  const current = workbook.sheets[index].tabColor;
  if (!wanted && !current) return [];
  if (wanted && current && current.rgb === wanted.rgb && current.theme === undefined && current.indexed === undefined && current.tint === undefined && current.auto === undefined) return [];
  return [{ type: 'setTabColor', sheet: index, color: wanted }];
}

export interface CopySheetOptions {
  /** Where the copy goes (0: first; the number of sheets: last). Default: right behind the sheet it is a copy of. */
  at?: number;
  /** The name of the copy. Default: copySheetName(). */
  name?: string;
}

/**
 * The name a copy of the sheet at `index` gets when none is given, as in Excel: the name of
 * the sheet with " (2)" behind it, or the lowest number from 2 that gives a name no sheet
 * has. A name too long for the 31 characters a sheet name has is cut to make room.
 */
export function copySheetName(workbook: Workbook, index: number): string {
  const sheet = workbook.sheets[index];
  if (!sheet) throw new RangeError(`No sheet at index ${index}`);
  const taken = new Set<string>();
  for (const other of workbook.sheets) taken.add(foldSheetName(other.name));
  // The name of a copy of a copy counts on: "Ad (2)" gives "Ad (3)".
  const numbered = /^(.*) \((\d{1,6})\)$/.exec(sheet.name);
  const base = numbered && numbered[1] !== '' ? numbered[1] : sheet.name;
  for (let n = 2; n <= MAX_DEFAULT_NUMBER; n++) {
    const tail = ` (${n})`;
    let head = base.slice(0, 31 - tail.length);
    // Not half of a character that takes two code units, and no name that ends up being none.
    if (/[\ud800-\udbff]$/.test(head)) head = head.slice(0, -1);
    const name = head + tail;
    if (taken.has(foldSheetName(name))) continue;
    try {
      checkSheetName(name);
    } catch (error) {
      continue;
    }
    return name;
  }
  throw new RangeError('The workbook has too many sheets');
}

/** Why the sheet at `index` cannot be copied, as a message; undefined when it can. */
export function canCopySheet(workbook: Workbook, index: number, options: CopySheetOptions = {}): string | undefined {
  return copySheetProblem(workbook, index, options.at, options.name);
}

/** What a copy of a sheet does not get: the kind of content and how much of it. */
export interface CopySheetLoss {
  /**
   * "conditionalFormatExtensions" (conditional rules the extension list has, or says more
   * about than the main part: those parts are not copied), "drawings" (pictures, shapes, charts),
   * "threadedComments", "externalLinks" (links to places outside the workbook), "pageBreaks",
   * "rowAttributes" (rows with attributes of another namespace than the sheet's own, the hint
   * x14ac:dyDescent aside; attributes without a namespace are copied), the kind of a part
   * around the sheet that the model does not hold ("table", "pivotTable", "formControl", ...),
   * or the name of something in the sheet part itself that the model does not hold
   * (Sheet.unmodelled: "sortState", "customSheetViews", "ignoredErrors", "cellWatches",
   * "sparklineGroups", "outlinePr", "codeName", "outlineLevels", "legacyDrawingHF" (the pictures
   * of headers and footers), "picture" (the background picture), ...).
   */
  what: string;
  count: number;
}

/** Kinds of parts around a sheet that the model holds, or that are counted by what the model has of them. */
const COUNTED_PARTS = ['drawing', 'image', 'comments', 'vmlDrawing', 'threadedComment', 'chart'];

/**
 * What copySheet() leaves behind of the sheet at `index`: content that is kept the way the
 * file has it and is written back from there, which a sheet that is in no file cannot be.
 * Empty when the copy holds all of the sheet.
 */
export function copySheetLosses(workbook: Workbook, index: number): CopySheetLoss[] {
  const sheet = workbook.sheets[index];
  if (!sheet) throw new RangeError(`No sheet at index ${index}`);
  const out: CopySheetLoss[] = [];
  const add = (what: string, count: number) => {
    if (count > 0) out.push({ what, count });
  };
  // The rules are copied; what the extension list says about them (a rule of its own, the border of a data bar) is not.
  add('conditionalFormatExtensions', extensionRuleCount(workbook, sheet));
  add('drawings', sheet.drawings ? sheet.drawings.length : 0);
  let threaded = 0;
  for (const comment of sheet.comments || []) if (comment && comment.threaded && commentCellNow(comment, sheet.structuralEdits)) threaded++;
  add('threadedComments', threaded);
  add('externalLinks', sheet.hyperlinks.filter((link) => link.location === undefined).length);
  const setup = sheet.pageSetup;
  if (setup && sheet.structuralEdits && sheet.structuralEdits.length > 0) add('pageBreaks', (setup.rowBreaks ? setup.rowBreaks.length : 0) + (setup.colBreaks ? setup.colBreaks.length : 0));
  for (const kind of sheet.anchored || []) if (COUNTED_PARTS.indexOf(kind) < 0 && !out.some((loss) => loss.what === kind)) add(kind, 1);
  // What the sheet part holds beside the fields of the model, by the names the reader noted.
  const names = new Map<string, number>();
  for (const name of sheet.unmodelled || []) if (typeof name === 'string') names.set(name, (names.get(name) || 0) + 1);
  for (const [name, count] of names) if (!out.some((loss) => loss.what === name)) add(name, count);
  // Attributes of rows that are kept as the file has them: those of a namespace of their own are not copied (the
  // new part does not declare it). The hints Excel puts on every row are no content.
  let rowAttributes = 0;
  const lost = new Map<NonNullable<Row['attrs']>, boolean>();
  for (const row of sheet.rows.values()) {
    const attrs = row.attrs;
    if (!attrs || attrs.length === 0) continue;
    // Equal lists are one object in a sheet as it was read: each is looked at once.
    let loses = lost.get(attrs);
    if (loses === undefined) {
      loses = attrs.some((pair) => !portableRowAttr(pair) && !isRowHint(pair));
      lost.set(attrs, loses);
    }
    if (loses) rowAttributes++;
  }
  add('rowAttributes', rowAttributes);
  return out;
}

/** Can a kept attribute of a row be written into a sheet part that declares no namespace but its own. */
function portableRowAttr(pair: [name: string, value: string]): boolean {
  return Array.isArray(pair) && typeof pair[0] === 'string' && typeof pair[1] === 'string' && /^[A-Za-z_][A-Za-z0-9_.-]*$/.test(pair[0]) && pair[0] !== 'xmlns';
}

/**
 * Is a kept attribute of a row a hint that says nothing about the sheet: the descent of the
 * font of the row (x14ac:dyDescent), which Excel writes on every row and works out again.
 */
function isRowHint(pair: [name: string, value: string]): boolean {
  const name = Array.isArray(pair) && typeof pair[0] === 'string' ? pair[0] : '';
  return name.slice(name.indexOf(':') + 1) === 'dyDescent' && name.indexOf(':') > 0;
}

/**
 * Copies the sheet at `index` into a new sheet of the same workbook, by default right behind
 * it and under the name copySheetName() gives. The copy holds what the model holds of the
 * sheet as it is now: cells with their values, formats and formulas, the sizes and formats of
 * rows and columns, merged cells, links to places in the workbook, the filter, how the sheet
 * is shown (frozen panes, zoom) and printed, the colour of its tab, its notes, its data
 * validations and conditional formats (as they apply now), and the defined names that are
 * local to it (print area, print titles, the range of its filter).
 * A reference without a sheet name is the same text in the copy and reads the copy; one that
 * names the copied sheet names the copy (UNSURE: so does the end of a 3D reference that is
 * the copied sheet); references to other sheets stay. Stored results of formulas are copied
 * and nothing is calculated.
 *
 * What is kept the way the file has it is not copied: see copySheetLosses(), which the caller
 * asks to tell the user. Nothing is refused for it: a sheet with a table or a pivot table is
 * copied as the cells it has. Throws RangeError for a sheet that is no worksheet, an index
 * outside the workbook and a name that cannot be used.
 */
export function copySheet(workbook: Workbook, index: number, options: CopySheetOptions = {}): Op[] {
  const at = options.at !== undefined ? options.at : index + 1;
  const problem = copySheetProblem(workbook, index, at, options.name);
  if (problem) throw new RangeError(problem);
  const source = workbook.sheets[index];
  const name = options.name !== undefined ? options.name : copySheetName(workbook, index);
  let sheetId = 0;
  for (const other of workbook.sheets) if (other.sheetId > sheetId) sheetId = other.sheetId;
  const copy = createSheet({ name, sheetId: sheetId + 1, state: 'visible', kind: 'worksheet', part: '' });

  // The rows and columns moved since the sheet was read: what the model keeps as the file has it is read through them.
  const space = ruleSpace(workbook, source);
  const own = foldSheetName(source.name);
  const map: SheetNameMap = (written) => (foldSheetName(written) === own ? name : undefined);
  const renamed = (text: string): string => (text.indexOf('!') < 0 ? text : renameSheetsInFormula(text, map));

  /** The attributes each list of kept row attributes gives the copy: equal lists of the file are one object, and stay one. */
  const rowAttrs = new Map<NonNullable<Row['attrs']>, NonNullable<Row['attrs']>>();
  // Cells are never altered in place: the copy shares every cell whose formula does not name the sheet.
  for (const [rowIndex, row] of source.rows) {
    if (!(rowIndex >= 0 && rowIndex < MAX_ROWS)) continue;
    let cells: Cell[] | undefined;
    row.cells.forEach((cell, col) => {
      const formula = cell.formula;
      if (!formula || formula.text === '' || col >= MAX_COLS) return;
      const text = renamed(formula.text);
      if (text === formula.text) return;
      if (!cells) cells = row.cells.slice();
      cells[col] = { ...cell, formula: { ...formula, text } };
    });
    const next: Row = { ...row, cells: cells || row.cells.slice() };
    // Attributes kept from the file: those of a namespace of their own stay behind (the new part does not declare it).
    if (row.attrs) {
      let attrs = rowAttrs.get(row.attrs);
      if (!attrs) rowAttrs.set(row.attrs, (attrs = row.attrs.filter(portableRowAttr)));
      if (attrs.length > 0) next.attrs = attrs;
      else delete next.attrs;
    }
    copy.rows.set(rowIndex, next);
  }
  copy.rowCount = source.rowCount;
  copy.colCount = source.colCount;
  copy.columns = source.columns.map((range) => ({ ...range }));
  copy.merges = source.merges.map((range) => ({ ...range }));
  const links: Hyperlink[] = [];
  for (const link of source.hyperlinks) {
    if (link.location === undefined) continue;
    const kept: Hyperlink = { range: { ...link.range }, location: renamed(link.location) };
    if (link.display !== undefined) kept.display = link.display;
    if (link.tooltip !== undefined) kept.tooltip = link.tooltip;
    links.push(kept);
  }
  copy.hyperlinks = links;
  if (source.defaultRowHeight !== undefined) copy.defaultRowHeight = source.defaultRowHeight;
  if (source.defaultColWidth !== undefined) copy.defaultColWidth = source.defaultColWidth;
  if (source.baseColWidth !== undefined) copy.baseColWidth = source.baseColWidth;
  copy.view = { ...source.view, tabSelected: false };
  if (source.tabColor) copy.tabColor = { ...source.tabColor };
  // A protected sheet has a protected copy, with the ranges that are open in it where they are now.
  if (source.protection) copy.protection = protectionNow(source.protection, source.structuralEdits, source.name);
  if (source.autoFilter) {
    // What the file has beside the fields of the model (ids, extensions) belongs to the filter of the file.
    const filter = { ...source.autoFilter };
    delete filter.attrs;
    delete filter.extra;
    copy.autoFilter = filter;
  }
  if (source.pageSetup) {
    const setup = { ...source.pageSetup };
    // Page breaks are kept as the file has them; once rows or columns moved they are not where the file says.
    if (source.structuralEdits && source.structuralEdits.length > 0) {
      delete setup.rowBreaks;
      delete setup.colBreaks;
    }
    copy.pageSetup = setup;
  }
  const notes: SheetComment[] = [];
  const seen = new Set<number>();
  for (const comment of source.comments || []) {
    if (!comment || !(comment.row >= 0 && comment.row < MAX_ROWS && comment.col >= 0 && comment.col < MAX_COLS)) continue;
    const cell = commentCellNow(comment, source.structuralEdits);
    if (!cell || seen.has(cell.row * MAX_COLS + cell.col)) continue;
    seen.add(cell.row * MAX_COLS + cell.col);
    if (comment.threaded) continue;
    const note: SheetComment = { row: cell.row, col: cell.col, text: String(comment.text), since: 0 };
    if (comment.author !== undefined) note.author = comment.author;
    notes.push(note);
  }
  if (notes.length > 0) copy.comments = notes;
  // The data validations, as they apply now: in the copy every one of them is a rule of its own (live),
  // those of the extension list included (their source is named in the rule itself from here on).
  const validations: DataValidation[] = [];
  for (const rule of source.dataValidations || []) {
    const now = dataValidationNow(rule, space);
    if (!now) continue;
    const copied: DataValidation = { ...now, ranges: now.ranges.map((range) => ({ ...range })), live: true };
    delete copied.src;
    delete copied.extended;
    for (const key of ['formula1', 'formula2'] as const) if (typeof copied[key] === 'string') copied[key] = renamed(copied[key] as string);
    validations.push(copied);
  }
  if (validations.length > 0) copy.dataValidations = validations;
  // The conditional formats, as they apply now, as far as the main part of a sheet states them (see copySheetLosses()).
  const formats: ConditionalFormat[] = [];
  for (const format of source.conditionalFormats || []) {
    const now = conditionalFormatNow(format, space);
    if (!now) continue;
    const rules: ConditionalRule[] = [];
    for (const rule of now.rules) {
      const kept = rule ? mainPartRule(rule) : undefined;
      if (!kept) continue;
      // Every formula of the rule, those of its thresholds included, names the copy where it named the sheet.
      rules.push(ruleWithFormulas(kept, renamed));
    }
    if (rules.length > 0) formats.push({ ranges: now.ranges.map((range) => ({ ...range })), rules, live: true });
  }
  if (formats.length > 0) copy.conditionalFormats = formats;
  copy.modified = true;

  const names: DefinedName[] = [];
  for (const defined of workbook.definedNames) if (defined.localSheet === index) names.push({ ...defined, value: renamed(defined.value) });
  return [{ type: 'copySheet', sheet: at, copy, names }];
}
