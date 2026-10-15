// Making a pivot table: from a range of cells whose first row names the fields (or a defined name
// that stands for one), on a new sheet or at a cell of a worksheet.
//
// A table is made the way a table of a file is refreshed: the definitions of a new table and of its
// cache are made with nothing in them (src/formats/xlsx/pivot-create.ts), read into the model by
// the reader as the parts of a file are, and the refresh builds the table in the cells and fills
// in the items, the records and the lines (src/ops/pivots.ts). What a refresh refuses, making a
// table refuses. The steps are one undo step: a new sheet ("copySheet" of an empty sheet that is in
// no file), the table and its cache ("addPivot"), and the refresh ("setPivot", "setCells").

import { readMadePivot } from '../formats/xlsx/pivot';
import { madeCacheXml, madeRecordsXml, madeTableXml, type MadeSource } from '../formats/xlsx/pivot-create';
import { MAX_PIVOT_SOURCE_CELLS, PIVOT_AGGREGATES, defaultValueName, namedSourceRange, sourceArea, type PivotAggregate } from '../model/sheet/pivot';
import { resolveCellStyle } from '../model/sheet/color';
import { MAX_COLS, MAX_ROWS, cellName, createSheet, foldSheetName, getCell, parseRange, rangeName, type CellRange, type Sheet, type Workbook } from '../model/sheet/types';
import { editsOf, restoreEdits } from '../model/edits';
import { applyOp, ensureCellFormat, ensureNumberFormat, markStyles, takeAddedStyles, type Op } from './ops';
import { PivotRefreshError, refreshOps, sourceColumns } from './pivots';
import { layoutProblem } from './pivot-fields';
import { tableArea } from './pivot-places';
import { currentRegion } from './sort';
import { defaultSheetName } from './sheets';

/** Where a table is made from: a range of a worksheet (its first row names the fields), or a defined name that stands for one. */
export type PivotSourceSpec = { sheet: number; range: CellRange } | { name: string };

/** Where a table goes: a new sheet, or a cell of a worksheet (the top left of the table, its filters included). */
export type PivotTargetSpec = { newSheet: true; name?: string; at?: number } | { sheet: number; row: number; col: number };

/** A table to make. Fields are the columns of the source, from 0. What is left out is chosen as Excel would suggest it. */
export interface PivotCreateOptions {
  source: PivotSourceSpec;
  /** By default a new sheet after the sheet of the source. */
  target?: PivotTargetSpec;
  rows?: number[];
  cols?: number[];
  filters?: number[];
  values?: { field: number; subtotal?: PivotAggregate; numFmt?: string }[];
  /** The name of the table; by default "PivotTable" with the first number no table of the workbook has. */
  name?: string;
  /** The language the name of a new sheet is made in (see defaultSheetName()). */
  locale?: string;
}

/** What making a table comes to: the operations (one undo step), and where the table will be. */
export interface PivotCreation {
  ops: Op[];
  /** Its place in Workbook.pivots.tables once the operations are applied. */
  index: number;
  /** The sheet it stands on (its index once the operations are applied). */
  sheet: number;
  /** Where its body will stand. */
  range: CellRange;
}

/** A table that is not made, with the reason. */
export class PivotCreateError extends RangeError {
  constructor(reason: string) {
    super(`Cannot make the pivot table: ${reason}`);
    this.name = 'PivotCreateError';
  }
}

const refuse = (reason: string): never => {
  throw new PivotCreateError(reason);
};

/** Number formats the file names by their number (built into the application), by their code. */
const BUILTIN_CODES: Record<string, number> = { '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, '0.00E+00': 11 };

/** The cells and the sheet a source stands for, and how a cache names it. */
interface Source {
  sheet: Sheet;
  range: CellRange;
  made: MadeSource;
}

/** The sheet and the cells of a source; refuses one that is not a range of a worksheet with a header and rows below it. */
function sourceOf(workbook: Workbook, spec: PivotSourceSpec): Source {
  if (!spec || typeof spec !== 'object') return refuse('its source is not given');
  let sheet: Sheet | undefined;
  let range: CellRange | undefined;
  let made: MadeSource;
  if ('name' in spec) {
    if (typeof spec.name !== 'string' || spec.name === '') return refuse('its source is not given');
    const found = namedSourceRange(workbook, { name: spec.name });
    if (typeof found === 'string') {
      // A name of whole columns or rows ("Data!$A:$C") stands for the cells of them that hold something, as such a range does.
      const whole = wholeNamed(workbook, spec.name);
      if (!whole) return refuse(found);
      return sourceOf(workbook, whole);
    }
    const wanted = foldSheetName(found.sheetName);
    sheet = workbook.sheets.find((one) => foldSheetName(one.name) === wanted);
    range = parseRange(found.ref.replace(/\$/g, ''));
    made = { kind: 'name', name: spec.name };
  } else {
    sheet = typeof spec.sheet === 'number' && Math.floor(spec.sheet) === spec.sheet ? workbook.sheets[spec.sheet] : undefined;
    const given = spec.range;
    if (!sheet) return refuse(`there is no sheet ${String(spec.sheet).slice(0, 20)}`);
    if (!given || typeof given !== 'object' || ![given.top, given.left, given.bottom, given.right].every((n) => typeof n === 'number' && n >= 0 && Math.floor(n) === n)) return refuse('its source is not a range of cells');
    // One cell stands for the block of cells around it, as in Excel; whole rows or columns for the part of them that holds cells.
    range =
      given.top === given.bottom && given.left === given.right
        ? currentRegion(sheet, given.top, given.left)
        : { top: given.top, left: given.left, bottom: Math.min(given.bottom, Math.max(given.top, sheet.rowCount - 1), MAX_ROWS - 1), right: Math.min(given.right, Math.max(given.left, sheet.colCount - 1), MAX_COLS - 1) };
    made = { kind: 'range', sheet: sheet.name, ref: rangeName(range) };
  }
  if (!sheet || sheet.kind !== 'worksheet') return refuse('its source is not on a worksheet');
  if (!range || range.bottom < range.top || range.right < range.left) return refuse('its source is not a range of cells');
  if (range.bottom <= range.top) return refuse('its source has no rows below its header');
  if ((range.bottom - range.top) * (range.right - range.left + 1) > MAX_PIVOT_SOURCE_CELLS) return refuse(`its source has more than ${MAX_PIVOT_SOURCE_CELLS} cells`);
  return { sheet, range, made };
}

/** A defined name of the workbook that stands for whole columns or rows of one sheet, as a range of that sheet; undefined for any other. */
function wholeNamed(workbook: Workbook, name: string): { sheet: number; range: CellRange } | undefined {
  const wanted = name.toLowerCase();
  const found = workbook.definedNames.find((one) => one.name.toLowerCase() === wanted && one.localSheet === undefined);
  const match = found ? /^(?:'((?:[^']|'')+)'|([^!'"\s]+))!(\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d{1,7}:\$?\d{1,7})$/.exec(found.value.trim()) : null;
  if (!match) return undefined;
  const sheetName = foldSheetName(match[1] !== undefined ? match[1].replace(/''/g, "'") : match[2]);
  const sheet = workbook.sheets.findIndex((one) => foldSheetName(one.name) === sheetName);
  const range = sourceArea(match[3]);
  return sheet >= 0 && range ? { sheet, range } : undefined;
}

/** The first name "PivotTable<n>" no table of the workbook has. */
function freeTableName(workbook: Workbook): string {
  const taken = new Set((workbook.pivots ? workbook.pivots.tables : []).map((table) => table.name.toLowerCase()));
  for (let n = 1; ; n++) if (!taken.has(`pivottable${n}`)) return `PivotTable${n}`;
}

const isIndex = (value: unknown, count: number): value is number => typeof value === 'number' && value >= 0 && value < count && Math.floor(value) === value;

/** The fields Excel's dialog would put first: the first column that is not all numbers in the rows, the first column of numbers (another one) summed. */
function suggested(numeric: boolean[]): { rows: number[]; values: { field: number; subtotal: PivotAggregate }[] } {
  const label = numeric.findIndex((plain) => !plain);
  const row = label >= 0 ? label : 0;
  const number = numeric.findIndex((plain, f) => plain && f !== row);
  return { rows: [row], values: number >= 0 ? [{ field: number, subtotal: 'sum' }] : [{ field: row, subtotal: 'count' }] };
}

/** The fields of a source as a table made from it would have them: their names, and whether each holds numbers only. */
export function pivotSourceFields(workbook: Workbook, source: PivotSourceSpec): { names: string[]; numeric: boolean[]; range: CellRange; sheet: number } {
  const found = sourceOf(workbook, source);
  const columns = sourceColumns(workbook, found.sheet, found.range);
  return { names: columns.names, numeric: columns.numeric, range: found.range, sheet: workbook.sheets.indexOf(found.sheet) };
}

/**
 * The operations that make a pivot table (see PivotCreateOptions), as one undo step. Throws
 * PivotCreateError (a RangeError) with the reason when it is not made: a source without a header
 * of names (an empty one, or one two columns share), a layout that is not built here, a place that
 * holds something or overlaps the source, a field with too many items for the rows, what a refresh
 * refuses. The workbook is as it was when this returns.
 */
export function createPivot(workbook: Workbook, options: PivotCreateOptions): PivotCreation {
  // The formats the table needs are added while it is built, and taken out again: they come back as the first step, which is undone with the rest.
  const mark = markStyles(workbook.styles);
  let made: PivotCreation | undefined;
  try {
    made = makePivot(workbook, options);
  } finally {
    const again = takeAddedStyles(workbook.styles, mark, made ? made.sheet : 0);
    if (made && again) made.ops.unshift(again);
  }
  return made;
}

function makePivot(workbook: Workbook, options: PivotCreateOptions): PivotCreation {
  if (!options || typeof options !== 'object') return refuse('nothing is given');
  // Pivot tables of the file that the model could not read: their caches and places would not be known.
  if (!workbook.pivots && workbook.sheets.some((sheet) => (sheet.anchored || []).indexOf('pivotTable') >= 0)) return refuse('the pivot tables of the workbook cannot be read');
  const source = sourceOf(workbook, options.source);
  const columns = sourceColumns(workbook, source.sheet, source.range);
  const names = columns.names;
  const seen = new Set<string>();
  names.forEach((name, c) => {
    if (name === '') refuse(`the header of column ${cellName(source.range.top, source.range.left + c)} of its source is empty`);
    if (seen.has(name.toLowerCase())) refuse(`two columns of its source have the header "${name.slice(0, 60)}"`);
    seen.add(name.toLowerCase());
  });
  const count = names.length;
  const guess = suggested(columns.numeric);
  const rows = options.rows !== undefined ? options.rows : guess.rows;
  const cols = options.cols !== undefined ? options.cols : [];
  const filters = options.filters !== undefined ? options.filters : [];
  const values: { field: number; subtotal?: PivotAggregate; numFmt?: string }[] = options.values !== undefined ? options.values : guess.values;
  const wrong = layoutProblem(count, rows, cols, filters, values);
  if (wrong !== undefined) refuse(wrong);
  const styles = workbook.styles;
  const taken = new Set<string>();
  const data = values.map((value) => {
    if (!value || typeof value !== 'object' || !isIndex(value.field, count)) return refuse('a value names a field the source does not have');
    const subtotal = value.subtotal ?? 'sum';
    if ((PIVOT_AGGREGATES as readonly string[]).indexOf(subtotal) < 0) refuse(`a value is summarized by a function that is not known here ("${String(subtotal).slice(0, 20)}")`);
    let name = defaultValueName(subtotal, names[value.field]);
    for (let n = 2; taken.has(name.toLowerCase()) || seen.has(name.toLowerCase()); n++) name = `${defaultValueName(subtotal, names[value.field])}${n}`;
    taken.add(name.toLowerCase());
    let numFmtId: number | undefined;
    if (value.numFmt !== undefined) {
      if (typeof value.numFmt !== 'string' || value.numFmt.length > 255) refuse('a number format is not one');
      const code = value.numFmt.trim();
      if (code !== '' && !/^general$/i.test(code)) numFmtId = BUILTIN_CODES[code] ?? ensureNumberFormat(styles, code);
    }
    return { fld: value.field, subtotal, name, numFmtId };
  });
  const name = options.name !== undefined ? options.name : freeTableName(workbook);
  if (typeof name !== 'string' || name.trim() === '' || name.length > 255) refuse('its name is not one');
  if ((workbook.pivots ? workbook.pivots.tables : []).some((table) => table.name.toLowerCase() === name.toLowerCase())) refuse(`a pivot table of the workbook is called "${name.slice(0, 60)}" already`);

  // Where it goes. A new sheet has its table at A3 (two rows above it for filters), as Excel puts it.
  const target: PivotTargetSpec = options.target || { newSheet: true };
  const above = filters.length > 0 ? filters.length + 1 : 0;
  let sheet: Sheet;
  let sheetIndex: number;
  let newSheet: Sheet | undefined;
  let at: { row: number; col: number };
  if ('newSheet' in target) {
    const sheetName = target.name !== undefined ? target.name : defaultSheetName(workbook, options.locale);
    const place = target.at !== undefined ? target.at : workbook.sheets.indexOf(source.sheet) + 1;
    if (!isIndex(place, workbook.sheets.length + 1)) refuse(`a sheet cannot be put at ${String(place).slice(0, 20)}`);
    let sheetId = 0;
    for (const other of workbook.sheets) if (other.sheetId > sheetId) sheetId = other.sheetId;
    newSheet = createSheet({ name: sheetName, sheetId: sheetId + 1, state: 'visible', kind: 'worksheet', part: '' });
    newSheet.modified = true;
    sheet = newSheet;
    sheetIndex = place;
    at = { row: Math.max(2, above), col: 0 };
  } else {
    const found = isIndex(target.sheet, workbook.sheets.length) ? workbook.sheets[target.sheet] : undefined;
    if (!found || found.kind !== 'worksheet') return refuse('its place is not on a worksheet');
    if (!isIndex(target.row, MAX_ROWS) || !isIndex(target.col, MAX_COLS)) return refuse('its place is not a cell');
    sheet = found;
    sheetIndex = target.sheet;
    at = { row: target.row + above, col: target.col };
    if (at.row >= MAX_ROWS) refuse('it would reach past the end of the sheet');
  }
  const spec = { name, fields: names, rows, cols, pages: filters, values: data, at, rowGrandTotals: true, colGrandTotals: true };
  // A column of dates gives its field the format of its dates, as Excel does.
  const formats = names.map((_, c) => {
    if (!columns.dates[c]) return 0;
    for (let r = source.range.top + 1; r <= source.range.bottom; r++) {
      const cell = getCell(source.sheet, r, source.range.left + c);
      if (cell && cell.kind === 'number') return resolveCellStyle(styles, cell.style).numFmtId;
    }
    return 0;
  });
  // A date that is a label is shown with that format (or a format of dates), and a value with its own: the style sheet gets a style for each.
  rows.concat(cols, filters).forEach((f) => {
    if (columns.dates[f]) ensureCellFormat(styles, 0, formats[f] || 14);
  });
  for (const value of data) if (value.numFmtId !== undefined) ensureCellFormat(styles, 0, value.numFmtId);
  const cacheXml = madeCacheXml(source.made, names, formats);
  const tableXml = madeTableXml(spec);
  let read: ReturnType<typeof readMadePivot>;
  try {
    read = readMadePivot(cacheXml, tableXml, sheet, workbook.date1904);
  } catch (error) {
    return refuse(error instanceof Error ? error.message.replace(/^Cannot make the pivot table: /, '') : String(error));
  }
  const { cache, table } = read;
  cache.created = { definition: cacheXml, records: madeRecordsXml() };
  table.created = tableXml;
  const index = workbook.pivots ? workbook.pivots.tables.length : 0;

  // The steps are applied here to build the table on the workbook as they leave it, and taken back before this returns.
  const steps: Op[] = [];
  if (newSheet) steps.push({ type: 'copySheet', sheet: sheetIndex, copy: newSheet, names: [] });
  steps.push({ type: 'addPivot', sheet: sheetIndex, cache, table });
  const edits = editsOf(workbook);
  const flags = { pivots: workbook.pivotsModified, sheets: workbook.sheetsModified };
  const undo: Op[] = [];
  let built: Op[];
  let range: CellRange;
  try {
    for (const op of steps) undo.push(applyOp(workbook, op));
    try {
      built = refreshOps(workbook, index, false);
    } catch (error) {
      if (error instanceof PivotRefreshError) refuse(error.message.replace(/^Cannot refresh the pivot table "[^"]*": /, ''));
      throw error;
    }
    const state = built.find((op) => op.type === 'setPivot');
    const tableState = state && state.type === 'setPivot' ? state.tables[0].state : undefined;
    if (!tableState) return refuse('it could not be built');
    range = tableState.location.ref;
    // The body and the filters above it: on cells that hold nothing, and not over the source.
    const area = tableArea(table, range);
    for (let r = area.top; r <= area.bottom; r++) {
      const row = sheet.rows.get(r);
      if (!row) continue;
      for (let c = area.left; c <= area.right; c++) {
        const cell = row.cells[c];
        if (cell && (cell.value !== null || cell.formula !== undefined)) refuse(`its place holds something (${cellName(r, c)})`);
      }
    }
    if (sheet === source.sheet && area.top <= source.range.bottom && area.bottom >= source.range.top && area.left <= source.range.right && area.right >= source.range.left) {
      refuse('it would stand over its source');
    }
    for (const merge of sheet.merges) if (merge.top <= area.bottom && merge.bottom >= area.top && merge.left <= area.right && merge.right >= area.left) refuse(`it would stand over the merged cells ${rangeName(merge)}`);
  } finally {
    for (let i = undo.length - 1; i >= 0; i--) applyOp(workbook, undo[i]);
    restoreEdits(workbook, edits);
    // Nothing was changed: the workbook says so as it did.
    if (flags.pivots) workbook.pivotsModified = true;
    else delete workbook.pivotsModified;
    if (flags.sheets) workbook.sheetsModified = true;
    else delete workbook.sheetsModified;
  }
  return { ops: steps.concat(built), index, sheet: sheetIndex, range };
}

/** Why a table would not be made; undefined when it would. Reads the source to tell. */
export function createPivotProblem(workbook: Workbook, options: PivotCreateOptions): string | undefined {
  try {
    createPivot(workbook, options);
    return undefined;
  } catch (error) {
    if (error instanceof RangeError) return error.message;
    throw error;
  }
}
