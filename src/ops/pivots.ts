// Refreshing pivot tables: the cache is read again from its source, and every table made from it
// is built again in the cells of its sheet, the way Excel builds it.
//
// Like the other commands these functions only build operations; the caller applies them
// through a History so that one refresh is one undo step. A refresh writes cells ("setCells":
// the labels, totals and values of the table, and empty cells where the table no longer
// reaches) and the state of the cache and its tables ("setPivot"), which the writer turns into
// the parts of the file.
//
// What is built here is a subset of what Excel builds, chosen so that what is built is what
// Excel shows: rows by one or more fields in compact, outline or tabular form, columns by one or
// more fields (compact form) or by the values, subtotals and grand totals, hidden items and
// fields of the page, eleven summary functions. A table with anything else (grouped items,
// calculated fields and items, filters by label or value, values shown as a share or a
// difference, data from outside the workbook, ...) is refused, and the reason is given: a
// table Excel would not make is never written.

import { resolveCellStyle } from '../model/sheet/color';
import {
  MAX_PIVOT_OUTPUT_CELLS,
  MAX_PIVOT_WORK,
  MAX_PIVOT_SOURCE_CELLS,
  cacheFieldsNow,
  cacheSourceNow,
  namedSourceRange,
  tableState,
  valueKey,
  type PivotAggregate,
  type PivotCache,
  type PivotCacheField,
  type PivotCacheState,
  type PivotItem,
  type PivotLine,
  type PivotLocation,
  type PivotTable,
  type PivotTableState,
  type PivotValue,
} from '../model/sheet/pivot';
import { MAX_COLS, MAX_ROWS, cellName, foldSheetName, getCell, parseRange, rangeName, type Cell, type CellRange, type Sheet, type Workbook } from '../model/sheet/types';
import { isoToSerial } from '../format/cell-format';
import { formatValue, isDateFormat } from '../format/number-format';
import { editsOf, readingsOf, restoreEdits } from '../model/edits';
import { applyOp, sameFormat, type CellChange, type Op } from './ops';

/** Counts for the tests: keys of text made to sort items. */
export const PIVOT_STATS = { textKeys: 0, scans: 0 };

/** Limits of a refresh that tests may lower to reach them with small files. */
export const PIVOT_LIMITS = { work: MAX_PIVOT_WORK };

/** A pivot table as a list of the workbook shows it. */
export interface PivotInfo {
  /** Its place in Workbook.pivots.tables: what the commands take. */
  index: number;
  name: string;
  /** The index of its sheet; -1 when the sheet is gone. */
  sheet: number;
  /** The cells of its body now. */
  range: CellRange;
  /** Why it cannot be refreshed, when that is known without reading its source. */
  refusal?: string;
}

/** A refresh that cannot be made, with the reason. */
export class PivotRefreshError extends RangeError {
  constructor(message: string) {
    super(message);
    this.name = 'PivotRefreshError';
  }
}

const refuse = (table: PivotTable | undefined, reason: string): never => {
  throw new PivotRefreshError(table ? `Cannot refresh the pivot table "${table.name}": ${reason}` : `Cannot refresh the pivot table: ${reason}`);
};

/** The pivot tables of a workbook. */
export function listPivots(workbook: Workbook): PivotInfo[] {
  const pivots = workbook.pivots;
  if (!pivots) return [];
  return pivots.tables.map((table, index) => {
    const info: PivotInfo = { index, name: table.name, sheet: workbook.sheets.indexOf(table.sheet), range: { ...tableState(table).location.ref } };
    const reason = table.cache.refusal || table.cache.referred || table.refusal;
    if (reason !== undefined) info.refusal = reason;
    return info;
  });
}

/** The count of edits of the workbook the last lookThroughPivots() looked at it with. */
const LOOKED = new WeakMap<object, string>();

/**
 * Looks through the workbook for what refers to the caches (see WorkbookPivots.lookThrough), so
 * that the refusals listPivots() gives are complete - for a menu that says what is refused. What
 * was found holds as long as the workbook is what it was then: the count of its edits (an edit,
 * an undo - a chart that comes back) and of the files written from it (the package) are the same.
 * Otherwise it is looked through again; a refresh goes by the same rule, so that a menu and the
 * refresh it leads to look once.
 */
export function lookThroughPivots(workbook: Workbook): void {
  const pivots = workbook.pivots;
  if (!pivots || !pivots.lookThrough) return;
  const now = `${editsOf(workbook)}:${readingsOf(workbook)}`;
  if (LOOKED.get(pivots) === now) return;
  PIVOT_STATS.scans++;
  pivots.lookThrough(workbook);
  LOOKED.set(pivots, now);
}

/** The pivot table whose body holds a cell of a sheet; undefined when none does. */
export function pivotAt(workbook: Workbook, sheetIndex: number, row: number, col: number): number | undefined {
  const pivots = workbook.pivots;
  const sheet = workbook.sheets[sheetIndex];
  if (!pivots || !sheet) return undefined;
  for (let i = 0; i < pivots.tables.length; i++) {
    const table = pivots.tables[i];
    if (table.sheet !== sheet) continue;
    const ref = tableState(table).location.ref;
    if (row >= ref.top && row <= ref.bottom && col >= ref.left && col <= ref.right) return i;
  }
  return undefined;
}

// ---------------------------------------------------------------------------------------------
// The source

function sheetNamed(workbook: Workbook, name: string): Sheet | undefined {
  const wanted = foldSheetName(name);
  for (const sheet of workbook.sheets) if (foldSheetName(sheet.name) === wanted) return sheet;
  return undefined;
}

/** The sheet and the range a cache reads: where its source stands now (rows and columns may have moved it). */
function sourceRange(workbook: Workbook, cache: PivotCache, table: PivotTable): { sheet: Sheet; range: CellRange } {
  const source = cacheSourceNow(cache);
  if (!source) return refuse(table, 'its source is not stated');
  let sheetName: string;
  let ref: string;
  if (source.kind === 'range') {
    sheetName = source.sheet;
    ref = source.ref;
  } else {
    const found = namedSourceRange(workbook, source);
    if (typeof found === 'string') return refuse(table, found);
    sheetName = found.sheetName;
    ref = found.ref;
  }
  const sheet = sheetNamed(workbook, sheetName);
  if (!sheet || sheet.kind !== 'worksheet') return refuse(table, `the sheet of its source ("${sheetName}") is not in the workbook`);
  const range = parseRange(ref.replace(/\$/g, ''));
  if (!range) return refuse(table, `its source ("${ref}") is not a range of cells`);
  return { sheet, range };
}

/** Whether a cell of a style shows its number as a date; asked once for each style. */
function dateStyles(workbook: Workbook): (style: number) => boolean {
  const known = new Map<number, boolean>();
  return (style) => {
    let value = known.get(style);
    if (value === undefined) {
      try {
        value = isDateFormat(resolveCellStyle(workbook.styles, style).numFmt);
      } catch (error) {
        value = false;
      }
      known.set(style, value);
    }
    return value;
  };
}

/** The value of a cell of the source as a cache holds it. */
function valueOf(workbook: Workbook, cell: Cell | undefined, isDate: (style: number) => boolean): PivotValue {
  if (!cell || cell.value === null || cell.kind === 'blank') return { t: 'm' };
  switch (cell.kind) {
    case 'number':
      return typeof cell.value === 'number' && isFinite(cell.value) ? (isDate(cell.style) ? { t: 'd', v: cell.value } : { t: 'n', v: cell.value }) : { t: 'm' };
    case 'boolean':
      return { t: 'b', v: cell.value === true };
    case 'error':
      return { t: 'e', v: String(cell.value) };
    case 'date': {
      const serial = isoToSerial(String(cell.value), workbook.date1904);
      return serial === undefined ? { t: 's', v: String(cell.value) } : { t: 'd', v: serial };
    }
    default:
      return { t: 's', v: String(cell.value) };
  }
}

/** The text a header cell names its field by. */
function headerText(value: PivotValue): string {
  switch (value.t) {
    case 's':
    case 'e':
      return value.v;
    case 'n':
    case 'd':
      return String(value.v);
    case 'b':
      return value.v ? 'TRUE' : 'FALSE';
    default:
      return '';
  }
}

/** The rows of the source below its header, after the header was checked against the fields of the cache. */
function readSource(workbook: Workbook, cache: PivotCache, table: PivotTable): PivotValue[][] {
  const { sheet, range } = sourceRange(workbook, cache, table);
  const fields = cacheFieldsNow(cache);
  const width = range.right - range.left + 1;
  // A column inserted inside the source is no field of the cache (see fieldColumnsNow()); a file read with more columns than
  // fields has the same. Excel adds a field for it; a refresh here does not, so it is refused rather than leave it out.
  if (width > fields.length) refuse(table, `its source has ${width} columns and its cache ${fields.length} fields: columns were inserted into it, and a refresh here adds no fields`);
  if (width !== fields.length) refuse(table, `its source has ${width} columns and its cache ${fields.length} fields`);
  if (range.bottom <= range.top) refuse(table, 'its source has no rows below its header');
  if ((range.bottom - range.top) * width > MAX_PIVOT_SOURCE_CELLS) refuse(table, `its source has more than ${MAX_PIVOT_SOURCE_CELLS} cells`);
  const isDate = dateStyles(workbook);
  for (let c = 0; c < width; c++) {
    const header = headerText(valueOf(workbook, getCell(sheet, range.top, range.left + c), isDate));
    if (header === '') refuse(table, `the header of column ${cellName(range.top, range.left + c)} of its source is empty`);
    if (header !== fields[c].name) refuse(table, `the header of ${cellName(range.top, range.left + c)} of its source is "${header.slice(0, 60)}", the field of the cache is "${fields[c].name.slice(0, 60)}"`);
  }
  const rows: PivotValue[][] = [];
  for (let r = range.top + 1; r <= range.bottom; r++) {
    const row = sheet.rows.get(r);
    const values: PivotValue[] = new Array(width);
    for (let c = 0; c < width; c++) values[c] = valueOf(workbook, row ? row.cells[range.left + c] : undefined, isDate);
    rows.push(values);
  }
  return rows;
}

// ---------------------------------------------------------------------------------------------
// The cache

/** Whether a field stands in the rows, the columns or the page of a table: its items are listed then. */
const onAxis = (table: PivotTable, field: number): boolean => {
  const axis = table.fields[field] ? table.fields[field].axis : undefined;
  return axis === 'axisRow' || axis === 'axisCol' || axis === 'axisPage';
};

/**
 * The cache as the source has it now. The items of a field that lists them keep the places they
 * had where their values are still in the source; new values follow, in the order they come.
 */
export function buildCache(cache: PivotCache, rows: PivotValue[][], tables: PivotTable[]): PivotCacheState {
  const before = cacheFieldsNow(cache);
  const fields: PivotCacheField[] = [];
  const records: (number | PivotValue)[][] = rows.map(() => new Array(before.length));
  for (let f = 0; f < before.length; f++) {
    const old = before[f];
    let plain = true;
    for (let r = 0; r < rows.length && plain; r++) {
      const t = rows[r][f].t;
      if (t === 's' || t === 'b' || t === 'e') plain = false;
    }
    const listed = old.items !== undefined || !plain || tables.some((table) => onAxis(table, f));
    if (!listed) {
      fields.push({ name: old.name });
      for (let r = 0; r < rows.length; r++) records[r][f] = rows[r][f];
      continue;
    }
    const present = new Map<string, number>();
    for (let r = 0; r < rows.length; r++) {
      const key = valueKey(rows[r][f]);
      if (!present.has(key)) present.set(key, r);
    }
    const items: PivotValue[] = [];
    const place = new Map<string, number>();
    for (const value of old.items || []) {
      const key = valueKey(value);
      if (present.has(key) && !place.has(key)) {
        place.set(key, items.length);
        items.push(value);
      }
    }
    for (let r = 0; r < rows.length; r++) {
      const value = rows[r][f];
      const key = valueKey(value);
      let at = place.get(key);
      if (at === undefined) {
        at = items.length;
        place.set(key, at);
        items.push(value);
      }
      records[r][f] = at;
    }
    fields.push({ name: old.name, items });
  }
  return { fields, records };
}

// ---------------------------------------------------------------------------------------------
// Summaries

/** What the values of one cell of a table add up to so far. */
class Acc {
  records = 0;
  /** Values that are not blank (what "count" counts). */
  filled = 0;
  nums = 0;
  sum = 0;
  product = 1;
  min = Infinity;
  max = -Infinity;
  mean = 0;
  m2 = 0;
  error: string | undefined;

  add(value: PivotValue): void {
    this.records++;
    if (value.t === 'm') return;
    this.filled++;
    if (value.t === 'e') {
      if (this.error === undefined) this.error = value.v;
      return;
    }
    if (value.t !== 'n' && value.t !== 'd') return;
    const v = value.v;
    this.nums++;
    this.sum += v;
    this.product *= v;
    if (v < this.min) this.min = v;
    if (v > this.max) this.max = v;
    const delta = v - this.mean;
    this.mean += delta / this.nums;
    this.m2 += delta * (v - this.mean);
  }

  /** The value of the cell; undefined for a cell that shows nothing. A number past the largest one is #NUM!, as Excel has it. */
  result(fn: PivotAggregate): number | string | undefined {
    const value = this.raw(fn);
    return typeof value === 'number' && !isFinite(value) ? '#NUM!' : value;
  }

  private raw(fn: PivotAggregate): number | string | undefined {
    if (this.records === 0) return undefined;
    if (fn === 'count') return this.filled === 0 ? undefined : this.filled;
    if (this.error !== undefined) return this.error;
    switch (fn) {
      case 'sum':
        return this.sum;
      case 'countNums':
        return this.nums;
      case 'average':
        return this.nums === 0 ? '#DIV/0!' : this.sum / this.nums;
      case 'max':
        return this.nums === 0 ? 0 : this.max;
      case 'min':
        return this.nums === 0 ? 0 : this.min;
      case 'product':
        return this.nums === 0 ? 0 : this.product;
      case 'var':
        return this.nums < 2 ? '#DIV/0!' : this.m2 / (this.nums - 1);
      case 'varp':
        return this.nums < 1 ? '#DIV/0!' : this.m2 / this.nums;
      case 'stdDev':
        return this.nums < 2 ? '#DIV/0!' : Math.sqrt(this.m2 / (this.nums - 1));
      case 'stdDevp':
        return this.nums < 1 ? '#DIV/0!' : Math.sqrt(this.m2 / this.nums);
    }
    return undefined;
  }
}

// ---------------------------------------------------------------------------------------------
// Layout

type Form = 'compact' | 'outline' | 'tabular';

/** A line of the rows or columns of a table as it is laid out. */
interface Line {
  kind: 'item' | 'sub' | 'grand' | 'data' | 'value';
  /** Places among the items of each field, from the first field of the direction on. */
  path: number[];
  /** The value a "value" line is of. */
  d: number;
}

/** What a table is laid out by: the same for the table as it was and as it is made. */
interface Shape {
  form: Form;
  /** Fields of the rows and of the columns (without the values). */
  rows: number[];
  cols: number[];
  /** Values: how many. */
  values: number;
  /** Columns of labels, and the rows above the first row of values. */
  labelCols: number;
  firstDataRow: number;
  firstHeaderRow: number;
}

/** How the rows of a table are laid out, or why that is not done here. */
function shapeOf(table: PivotTable): Shape {
  const values = table.dataFields.length;
  if (values === 0) refuse(table, 'it has no values');
  if (table.rowFields.indexOf(-2) >= 0) refuse(table, 'its values stand in rows');
  const rows = table.rowFields.slice();
  if (rows.length === 0) refuse(table, 'it has no fields in its rows');
  const at = table.colFields.indexOf(-2);
  const cols = table.colFields.filter((x) => x !== -2);
  if (values > 1 && at < 0) refuse(table, 'its values have no place in its columns');
  if (at >= 0 && at !== table.colFields.length - 1) refuse(table, 'its values do not stand after the fields of its columns');
  if (at >= 0 && values < 2) refuse(table, 'its values stand in its columns although there is one of them');
  if (cols.length > 0 && values > 1) refuse(table, 'it has several values as well as fields in its columns');
  for (const f of rows.concat(cols)) {
    const field = table.fields[f];
    if (!field) refuse(table, 'it names a field it does not have');
  }
  const formOf = (f: number): Form => {
    const field = table.fields[f];
    return !field.outline ? 'tabular' : field.compact ? 'compact' : 'outline';
  };
  // The form of the fields above the last one decides; with one field, its own.
  const deciding = rows.length > 1 ? rows.slice(0, -1) : rows;
  const form = formOf(deciding[0]);
  for (const f of deciding) if (formOf(f) !== form) refuse(table, 'its fields in rows are laid out in different forms');
  if (cols.length > 0 && form !== 'compact') refuse(table, 'it has fields in its columns and is not in compact form');
  const labelCols = form === 'compact' ? 1 : rows.length;
  const firstDataRow = cols.length === 0 ? 1 : cols.length + 1;
  const firstHeaderRow = cols.length === 0 && values > 1 ? 0 : 1;
  return { form, rows, cols, values, labelCols, firstDataRow, firstHeaderRow };
}

/** The fields of the rows and columns of a table, without the values. */
const shapeOfFields = (table: PivotTable): number[] => table.rowFields.concat(table.colFields).filter((f) => f >= 0);

/** The lines of the file (rowItems, colItems) as laid-out lines; undefined when they cannot be followed. */
function linesOfFile(list: PivotLine[], valueLines: boolean): Line[] | undefined {
  const out: Line[] = [];
  let previous: number[] = [];
  for (const line of list) {
    if (line.r > previous.length) return undefined;
    const path = previous.slice(0, line.r).concat(line.x);
    if (line.t === 'grand') out.push({ kind: 'grand', path: [], d: line.i });
    else if (line.t !== undefined) out.push({ kind: 'sub', path, d: line.i });
    else if (valueLines) out.push({ kind: 'value', path: [], d: path.length > 0 ? path[path.length - 1] : 0 });
    else if (path.length === 0) out.push({ kind: 'data', path, d: 0 });
    else out.push({ kind: 'item', path, d: line.i });
    previous = path;
  }
  return out;
}

/** The lines of the file a list of laid-out lines is written as. */
function fileLines(lines: Line[]): PivotLine[] {
  const out: PivotLine[] = [];
  let previous: number[] = [];
  for (const line of lines) {
    if (line.kind === 'data') {
      out.push({ r: 0, i: 0, x: [] });
      continue;
    }
    if (line.kind === 'value') {
      out.push({ r: 0, i: line.d, x: [line.d] });
      previous = [line.d];
      continue;
    }
    if (line.kind === 'grand') {
      out.push({ t: 'grand', r: 0, i: 0, x: [0] });
      previous = [];
      continue;
    }
    const path = line.path;
    let common = 0;
    while (common < previous.length && common < path.length && previous[common] === path[common]) common++;
    const r = Math.min(common, path.length - 1);
    const entry: PivotLine = { r, i: 0, x: path.slice(r) };
    if (line.kind === 'sub') entry.t = 'default';
    out.push(entry);
    previous = path;
  }
  return out;
}

/** What a cell of a table is: where it stands within the body and what it is for. */
interface Placed {
  row: number;
  col: number;
  role: string;
  /** For labels: the field and the item; for values: the line of the rows and of the columns, and the value. */
  field?: number;
  item?: number;
  rowLine?: number;
  colLine?: number;
  d?: number;
  /** A caption: which. */
  caption?: 'rowHeader' | 'colHeader' | 'grand' | 'data' | 'field';
}

/** Places every cell of a table laid out by `shape` with these lines. */
function place(table: PivotTable, shape: Shape, rowLines: Line[], colLines: Line[], emit: (cell: Placed) => void): void {
  const { form, rows, cols, labelCols, firstDataRow } = shape;
  const header = firstDataRow - 1;
  // The header of the labels.
  if (form === 'compact') emit({ row: header, col: 0, role: 'rowHeader', caption: 'rowHeader' });
  else for (let j = 0; j < rows.length; j++) emit({ row: header, col: j, role: 'fieldHeader', caption: 'field', field: rows[j] });
  if (cols.length === 0) {
    colLines.forEach((line, j) => emit({ row: 0, col: labelCols + j, role: 'dataName', caption: 'data', d: line.kind === 'value' ? line.d : 0 }));
  } else {
    emit({ row: 0, col: 0, role: 'dataName', caption: 'data', d: 0 });
    emit({ row: 0, col: labelCols, role: 'colHeader', caption: 'colHeader' });
    let previous: number[] = [];
    colLines.forEach((line, j) => {
      const col = labelCols + j;
      if (line.kind === 'grand') emit({ row: 1, col, role: 'colGrand', caption: 'grand' });
      else if (line.kind === 'sub') emit({ row: 1 + line.path.length - 1, col, role: `colSub${line.path.length - 1}`, field: cols[line.path.length - 1], item: line.path[line.path.length - 1] });
      else if (line.kind === 'item') {
        let common = 0;
        while (common < previous.length && common < line.path.length && previous[common] === line.path[common]) common++;
        for (let q = Math.min(common, line.path.length - 1); q < line.path.length; q++) emit({ row: 1 + q, col, role: `colLabel${q}`, field: cols[q], item: line.path[q] });
      }
      previous = line.kind === 'grand' ? [] : line.path;
    });
  }
  // The lines of the rows: their labels, then their values.
  let previous: number[] = [];
  rowLines.forEach((line, i) => {
    const row = firstDataRow + i;
    const depth = line.path.length - 1;
    if (line.kind === 'grand') emit({ row, col: 0, role: 'rowGrand', caption: 'grand' });
    else if (line.kind === 'sub') emit({ row, col: form === 'compact' ? 0 : depth, role: `rowSub${depth}`, field: rows[depth], item: line.path[depth] });
    else if (form === 'tabular') {
      let common = 0;
      while (common < previous.length && common < line.path.length && previous[common] === line.path[common]) common++;
      for (let q = Math.min(common, depth); q <= depth; q++) emit({ row, col: q, role: `rowLabel${q}`, field: rows[q], item: line.path[q] });
    } else emit({ row, col: form === 'compact' ? 0 : depth, role: `rowLabel${depth}`, field: rows[depth], item: line.path[depth] });
    previous = line.kind === 'grand' ? [] : line.path;
    // A line above others of the rows (compact and outline form) shows the subtotal of its item only when it is shown at the top.
    if (line.kind === 'item' && depth < rows.length - 1) {
      const field = table.fields[rows[depth]];
      if (!field.defaultSubtotal || !field.subtotalTop) return;
    }
    const total = line.kind !== 'item' || depth < rows.length - 1;
    colLines.forEach((colLine, j) => {
      const d = colLine.kind === 'value' ? colLine.d : 0;
      const role = line.kind === 'grand' || colLine.kind === 'grand' ? 'grandValue' : total || colLine.kind === 'sub' ? 'subValue' : 'value';
      emit({ row, col: labelCols + j, role, rowLine: i, colLine: j, d });
    });
  });
}

/**
 * The words around the label of an item in the label of its subtotal when the table shows none
 * to go by: from the word of its grand total - "Genel Toplam" is Turkish Excel's, whose subtotals
 * read "{item} Toplam" (UNSURE) - else English, "{item} Total".
 */
function subtotalWords(grand: string | undefined): { before: string; after: string } {
  if (grand !== undefined && lowerTr(grand.trim()).replace(/\s+/g, ' ') === 'genel toplam') return { before: '', after: ' Toplam' };
  return { before: '', after: WORDS.total };
}

/** Excel's words for what the file leaves to the application, in English. */
const WORDS = { rowHeader: 'Row Labels', colHeader: 'Column Labels', grand: 'Grand Total', total: ' Total', blank: '(blank)' };

/** What the cells of a table as it was say of how the table is written: styles, captions in the language of the file. */
interface Harvest {
  styles: Map<string, number>;
  captions: { rowHeader?: string; colHeader?: string; grand?: string; blank?: string };
  /** The words around the label of an item in the label of its subtotal ("{item} Total"). */
  total?: { before: string; after: string };
}

/** The text a label of an item shows (for the label of its subtotal). */
function labelText(workbook: Workbook, value: PivotValue): string {
  switch (value.t) {
    case 's':
    case 'e':
      return value.v;
    case 'n':
      return formatValue(value.v, 'General').text;
    case 'd':
      return formatValue(value.v, 'm/d/yyyy', { date1904: workbook.date1904 }).text;
    case 'b':
      return value.v ? 'TRUE' : 'FALSE';
    default:
      return '';
  }
}

/** The caption an item is shown with: the name the table gives it (`n`), else its value. */
function itemCaption(item: PivotItem | undefined): string | undefined {
  if (!item) return undefined;
  for (const [name, value] of item.attrs) if (name === 'n') return value;
  return undefined;
}

/** Reads the table as it stands in its sheet: the style of each kind of cell, and the captions of the file. */
function harvest(workbook: Workbook, table: PivotTable, shape: Shape, state: PivotTableState, values: PivotCacheField[]): Harvest {
  const out: Harvest = { styles: new Map(), captions: {} };
  const rowLines = linesOfFile(state.rowItems, false);
  const colLines = linesOfFile(state.colItems, shape.cols.length === 0 && shape.values > 1);
  if (!rowLines || !colLines) return out;
  // Lines that do not fit where the table stands say nothing about its cells: they are not followed (whatever they would cost).
  const area = state.location.ref;
  if (state.location.firstDataRow + rowLines.length !== area.bottom - area.top + 1 || state.location.firstDataCol + colLines.length !== area.right - area.left + 1) return out;
  const old: Shape = { ...shape, labelCols: state.location.firstDataCol, firstDataRow: state.location.firstDataRow };
  const ref = state.location.ref;
  place(table, old, rowLines, colLines, (cell) => {
    const row = ref.top + cell.row;
    const col = ref.left + cell.col;
    if (row > ref.bottom || col > ref.right) return;
    const found = getCell(table.sheet, row, col);
    if (!found) return;
    const key = cell.d !== undefined && cell.role.indexOf('alue') >= 0 ? `${cell.role}:${cell.d}` : cell.role;
    if (!out.styles.has(key)) out.styles.set(key, found.style);
    const text = found.kind === 'string' && typeof found.value === 'string' ? found.value : undefined;
    if (text === undefined) return;
    if (cell.caption === 'rowHeader' && out.captions.rowHeader === undefined) out.captions.rowHeader = text;
    if (cell.caption === 'colHeader' && out.captions.colHeader === undefined) out.captions.colHeader = text;
    if (cell.caption === 'grand' && out.captions.grand === undefined) out.captions.grand = text;
    if (cell.field === undefined || cell.item === undefined) return;
    const items = state.items[cell.field];
    const item = items ? items[cell.item] : undefined;
    const value = item && item.x !== undefined && values[cell.field] && values[cell.field].items ? (values[cell.field].items as PivotValue[])[item.x] : undefined;
    if (!value) return;
    if (value.t === 'm' && /Label/.test(cell.role) && out.captions.blank === undefined && itemCaption(item) === undefined) out.captions.blank = text;
    if (/Sub/.test(cell.role) && out.total === undefined) {
      const label = itemCaption(item) ?? (value.t === 'm' ? out.captions.blank ?? WORDS.blank : labelText(workbook, value));
      const at = label === '' ? -1 : text.indexOf(label);
      if (at >= 0) out.total = { before: text.slice(0, at), after: text.slice(at + label.length) };
    }
  });
  return out;
}

// ---------------------------------------------------------------------------------------------
// Items

/** The letters in the order of the Turkish alphabet, with q, w and x where the Latin alphabet has them. */
const ALPHABET = 'abcçdefgğhıijklmnoöpqrsştuüvwxyz';

/** A Latin letter and the mark that make a Turkish letter of it. */
const TURKISH_LETTERS: Record<string, string> = { 'c\u0327': 'ç', 'g\u0306': 'ğ', 'o\u0308': 'ö', 's\u0327': 'ş', 'u\u0308': 'ü' };

/** What a text weighs at each level of the comparison (see compareText()). */
interface TextKey {
  primary: number[];
  secondary: number[];
  tertiary: number[];
}

/** Classes of primary weights: signs and numerals of the ASCII range, Latin letters, every other character. */
const SIGNS = 0;
const LATIN = 0x200000;
const OTHER = 0x400000;

/** Combining marks: they weigh on the second level only, as the accent of the character before them. */
const MARK = /^\p{M}$/u;
const isMark = (code: number) => code === 0x34f || MARK.test(String.fromCodePoint(code));

/**
 * The weights of a text, without Intl, on three levels as a collation has them. First: the
 * letter, without its case and its accent - the Turkish alphabet (Ç Ğ I İ Ö Ş Ü are letters of
 * their own, not accented ones; I and ı are one letter, İ and i another), Latin letters with an
 * accent count as their letter; signs and numerals of the ASCII range come before the letters,
 * other scripts after them, each letter in its lower case. Second: the accents. Third: the case.
 */
function textKey(text: string): TextKey {
  PIVOT_STATS.textKeys++;
  const key: TextKey = { primary: [], secondary: [], tertiary: [] };
  const normal = text.normalize('NFC');
  for (let i = 0; i < normal.length; ) {
    const code = normal.codePointAt(i) as number;
    i += code > 0xffff ? 2 : 1;
    if (isMark(code)) {
      key.secondary.push(code);
      continue;
    }
    const ch = String.fromCodePoint(code);
    const lower = ch === 'I' ? 'ı' : ch === 'İ' ? 'i' : ch.toLowerCase();
    const upper = lower !== ch ? 1 : 0;
    const own = lower.length === 1 ? ALPHABET.indexOf(lower) : -1;
    if (own >= 0) {
      key.primary.push(LATIN + own);
      key.secondary.push(0);
      key.tertiary.push(upper);
      continue;
    }
    const parts = lower.normalize('NFD');
    const base = parts.codePointAt(0) as number;
    let marks = parts.slice(base > 0xffff ? 2 : 1);
    // A Latin letter whose marks include the one that makes a Turkish letter of it (ǖ: ü and a macron) is that letter; the other marks are its accent.
    let letter = base < 0x80 ? String.fromCharCode(base) : '';
    for (let k = 0; k < marks.length && letter !== ''; k++) {
      const made = TURKISH_LETTERS[letter + marks.charAt(k)];
      if (made !== undefined) {
        letter = made;
        marks = marks.slice(0, k) + marks.slice(k + 1);
        break;
      }
    }
    let accent = 0;
    for (let k = 0; k < marks.length; k++) accent = accent * 0x400 + marks.charCodeAt(k);
    const latin = letter !== '' ? ALPHABET.indexOf(letter) : -1;
    if (latin >= 0) key.primary.push(LATIN + latin);
    else key.primary.push(code < 0x80 ? SIGNS + code : OTHER + base);
    key.secondary.push(accent);
    key.tertiary.push(upper);
  }
  return key;
}

function compareLists(a: number[], b: number[]): number {
  const n = Math.min(a.length, b.length);
  for (let i = 0; i < n; i++) if (a[i] !== b[i]) return a[i] - b[i];
  return a.length - b.length;
}

/**
 * Compares two texts as items are sorted: by their letters, then their accents, then their case
 * (see textKey()); texts that weigh the same on every level by their code units, so that the
 * order is a total one.
 */
export function compareText(a: string, b: string): number {
  if (a === b) return 0;
  return compareKeys(textKey(a), textKey(b), a, b);
}

/** Compares texts by their keys (made once for each, see textKey()). */
function compareKeys(x: TextKey, y: TextKey, a: string, b: string): number {
  const order = compareLists(x.primary, y.primary) || compareLists(x.secondary, y.secondary) || compareLists(x.tertiary, y.tertiary);
  if (order !== 0) return order;
  return a === b ? 0 : a < b ? -1 : 1;
}

/** Full names of months and days (English and Turkish) that Excel's own lists sort by the calendar. */
const CALENDAR_NAMES = new Set(
  (
    'january february march april may june july august september october november december ' +
    'sunday monday tuesday wednesday thursday friday saturday ' +
    'ocak şubat mart nisan mayıs haziran temmuz ağustos eylül ekim kasım aralık ' +
    'pazar pazartesi salı çarşamba perşembe cuma cumartesi'
  ).split(' '),
);
/** Their short names: words of their own as well (ara, cum, sal, per, mar...), counted only in a field of nothing but short names. */
const CALENDAR_SHORT = new Set(
  (
    'jan feb mar apr jun jul aug sep oct nov dec sun mon tue wed thu fri sat ' +
    'oca şub nis haz tem ağu eyl eki kas ara paz pzt sal çar per cum cmt'
  ).split(' '),
);

/** Lower case as Turkish has it: I is ı, İ is i. */
const lowerTr = (text: string) => text.replace(/I/g, 'ı').replace(/İ/g, 'i').toLowerCase();

/** What an item is sorted by, made once for each item: its kind, and its number or the key of its text. */
interface SortKey {
  rank: number;
  value: PivotValue;
  text?: TextKey;
}

function sortKeyOf(value: PivotValue): SortKey {
  const rank = value.t === 'n' || value.t === 'd' ? 0 : value.t === 's' ? 1 : value.t === 'b' ? 2 : value.t === 'e' ? 3 : 4;
  return value.t === 's' ? { rank, value, text: textKey(value.v) } : { rank, value };
}

/** Sorts values as Excel sorts the items of a field: numbers and dates, then text, logical values, errors, blanks. */
function compareSortKeys(x: SortKey, y: SortKey): number {
  if (x.rank !== y.rank) return x.rank - y.rank;
  const a = x.value;
  const b = y.value;
  if ((a.t === 'n' || a.t === 'd') && (b.t === 'n' || b.t === 'd')) return a.v - b.v;
  if (a.t === 's' && b.t === 's') return compareKeys(x.text as TextKey, y.text as TextKey, a.v, b.v);
  if (a.t === 'b' && b.t === 'b') return Number(a.v) - Number(b.v);
  if (a.t === 'e' && b.t === 'e') return a.v < b.v ? -1 : a.v > b.v ? 1 : 0;
  return 0;
}

/** An item with another place in the cache: its attributes say so. */
function withX(item: PivotItem, x: number): PivotItem {
  const attrs = item.attrs.slice();
  const at = attrs.findIndex(([name]) => name === 'x');
  if (at >= 0) attrs[at] = ['x', String(x)];
  else attrs.push(['x', String(x)]);
  const out: PivotItem = { x, attrs };
  if (item.t !== undefined) out.t = item.t;
  if (item.h) out.h = true;
  return out;
}

/**
 * The items of each field of a table made from the cache as it is now: the items it had keep
 * their order, their names and whether they are hidden, as long as their values are still
 * there; the new ones follow (or all are sorted, for a field that is sorted).
 */
function itemsNow(table: PivotTable, state: PivotTableState, before: PivotCacheField[], after: PivotCacheField[]): (PivotItem[] | undefined)[] {
  return table.fields.map((field, f) => {
    const old = state.items[f];
    const values = after[f] ? after[f].items : undefined;
    const axis = field.axis === 'axisRow' || field.axis === 'axisCol' || field.axis === 'axisPage';
    if (!values) {
      if (axis) refuse(table, `the field "${after[f] ? after[f].name : f}" lists no items`);
      return old;
    }
    if (old === undefined && !axis) return undefined;
    const oldValues = before[f] ? before[f].items : undefined;
    const place = new Map<string, number>();
    values.forEach((value, x) => place.set(valueKey(value), x));
    const used = new Set<number>();
    const data: PivotItem[] = [];
    for (const item of old || []) {
      if (item.t !== undefined) continue;
      const value = oldValues && item.x !== undefined ? oldValues[item.x] : undefined;
      if (!value) refuse(table, 'the items of a field do not match its cache');
      const x = place.get(valueKey(value as PivotValue));
      if (x === undefined || used.has(x)) continue;
      used.add(x);
      data.push(withX(item, x));
    }
    for (let x = 0; x < values.length; x++) {
      if (used.has(x)) continue;
      // A field that hides what is new to it (hideNewItems) shows only the items it knew.
      if (field.hideNewItems && old !== undefined) data.push({ x, h: true, attrs: [['h', '1'], ['x', String(x)]] });
      else data.push({ x, attrs: [['x', String(x)]] });
    }
    if (field.sortType !== 'manual') {
      // Excel sorts names of months and days by the calendar (its own lists), whatever else the field holds; that is not done
      // here. A full name refuses; short names, which are words of their own as well, only in a field of nothing but them.
      if (table.customListSort) {
        // In lower case the Turkish way (PAZARTESİ) and the plain way (FRIDAY): either may be the name.
        const words = values.filter((value) => value.t !== 'm').map((value) => (value.t === 's' ? [lowerTr(value.v.trim()), value.v.trim().toLowerCase()] : undefined));
        const full = words.some((word) => word !== undefined && (CALENDAR_NAMES.has(word[0]) || CALENDAR_NAMES.has(word[1])));
        const short = words.length > 0 && words.every((word) => word !== undefined && (CALENDAR_SHORT.has(word[0]) || CALENDAR_SHORT.has(word[1])));
        if (full || short) refuse(table, 'it sorts a field whose items are names of months or days (Excel sorts them by the calendar)');
      }
      const sign = field.sortType === 'descending' ? -1 : 1;
      // The keys are made once for each item; stable: equal values keep their order.
      const order = data.map((item, i) => ({ item, i, key: sortKeyOf(values[item.x as number]) }));
      order.sort((a, b) => sign * compareSortKeys(a.key, b.key) || a.i - b.i);
      for (let i = 0; i < order.length; i++) data[i] = order[i].item;
    }
    const rest = (old || []).filter((item) => item.t !== undefined);
    if (axis && field.defaultSubtotal && !rest.some((item) => item.t === 'default')) rest.push({ t: 'default', attrs: [['t', 'default']] });
    return data.concat(rest);
  });
}

// ---------------------------------------------------------------------------------------------
// A table built again

/** A table made from a refreshed cache: its state and its cells, by place within its body. */
interface Built {
  table: PivotTable;
  sheet: number;
  state: PivotTableState;
  /** The cells of the body: row, column and the cell. */
  cells: { row: number; col: number; cell: Cell }[];
}

/** The style a cell of a kind gets: what that kind had in the table as it was, else the plain one. */
function styleFor(harvested: Harvest, role: string, d: number | undefined): number {
  const styles = harvested.styles;
  if (d !== undefined) {
    // A value takes the style of a value of its own field first: the same format of numbers, whatever line it stands on.
    for (const key of [`${role}:${d}`, `value:${d}`, `subValue:${d}`, `grandValue:${d}`]) {
      const found = styles.get(key);
      if (found !== undefined) return found;
    }
  }
  const kind = styles.get(role);
  if (kind !== undefined) return kind;
  // Labels of fields further in take the style of the first one.
  const base = role.replace(/\d+$/, '0');
  const first = styles.get(base);
  return first !== undefined ? first : 0;
}

/** The style of a cell that shows a date: `base` with a date format when the style sheet has one like it, else `base`. */
function dateStyle(workbook: Workbook, base: number): number {
  const styles = workbook.styles;
  const format = styles.cellXfs[base];
  if (!format) return base;
  try {
    if (isDateFormat(resolveCellStyle(styles, base).numFmt)) return base;
  } catch (error) {
    return base;
  }
  const wanted = { ...format, numFmtId: 14 };
  const found = styles.cellXfs.findIndex((other) => sameFormat(other, wanted));
  return found >= 0 ? found : base;
}

/** The style of a value: `base`, or (for a style without a format of numbers) one with the format the value names, when the style sheet has it. */
function valueStyle(workbook: Workbook, base: number, numFmtId: number | undefined, harvested: boolean): number {
  if (harvested || numFmtId === undefined || numFmtId === 0) return base;
  const styles = workbook.styles;
  const format = styles.cellXfs[base];
  if (!format) return base;
  const wanted = { ...format, numFmtId };
  const found = styles.cellXfs.findIndex((other) => sameFormat(other, wanted));
  return found >= 0 ? found : base;
}

function build(workbook: Workbook, table: PivotTable, sheetIndex: number, after: PivotCacheState, before: PivotCacheField[]): Built {
  const shape = shapeOf(table);
  const old = tableState(table);
  const was = old.location.ref;
  if ((was.bottom - was.top + 1) * (was.right - was.left + 1) > MAX_PIVOT_OUTPUT_CELLS) refuse(table, `it has more than ${MAX_PIVOT_OUTPUT_CELLS} cells`);
  const harvested = harvest(workbook, table, shape, old, before);
  const items = itemsNow(table, old, before, after.fields);
  const fields = after.fields;
  // A field of the rows or columns that hides what is new to it, and so would hide every item: what Excel does then is not known.
  for (const f of shapeOfFields(table)) {
    const list = items[f];
    if (!table.fields[f].hideNewItems || !list) continue;
    const data = list.filter((item) => item.t === undefined);
    if (data.length > 0 && data.every((item) => item.h)) refuse(table, `every item of the field "${fields[f] ? fields[f].name : f}" would be hidden`);
  }
  /** The place of a value of the cache among the items of the table, by field. */
  const placeOf: (Int32Array | undefined)[] = items.map((list, f) => {
    if (!list) return undefined;
    const values = fields[f].items;
    const out = new Int32Array(values ? values.length : 0).fill(-1);
    list.forEach((item, i) => {
      if (item.t === undefined && item.x !== undefined && item.x < out.length) out[item.x] = i;
    });
    return out;
  });
  // What the page shows: one item, or every item that is not hidden.
  const pageItems: (number | undefined)[] = [];
  const pages: { f: number; only?: number }[] = [];
  table.pageFields.forEach((page, p) => {
    const was = old.pageItems[p];
    if (was === undefined) {
      pageItems.push(undefined);
      pages.push({ f: page.fld });
      return;
    }
    const oldItems = old.items[page.fld];
    const oldItem = oldItems ? oldItems[was] : undefined;
    const oldValues = before[page.fld] ? before[page.fld].items : undefined;
    const value = oldItem && oldItem.x !== undefined && oldValues ? oldValues[oldItem.x] : undefined;
    const list = items[page.fld];
    const now = value && list ? list.findIndex((item) => item.t === undefined && item.x !== undefined && fields[page.fld].items && valueKey((fields[page.fld].items as PivotValue[])[item.x]) === valueKey(value)) : -1;
    if (now < 0) refuse(table, `the item the page shows for the field "${fields[page.fld] ? fields[page.fld].name : page.fld}" is not in its source any more`);
    pageItems.push(now);
    pages.push({ f: page.fld, only: now });
  });

  const { rows, cols, values: valueCount, form } = shape;
  const n = rows.length;
  const m = cols.length;
  const D = valueCount;
  /** For a record: the item of each field of the rows and the columns, or -1 for one that is filtered out. */
  const itemOf = (record: (number | PivotValue)[], f: number): number => {
    const at = record[f];
    const map = placeOf[f];
    if (typeof at !== 'number' || !map || at < 0 || at >= map.length) return -1;
    return map[at];
  };
  const accs = new Map<string, Acc[]>();
  const children = new Map<string, Set<number>>();
  const add = (key: string, record: (number | PivotValue)[]) => {
    let list = accs.get(key);
    if (!list) {
      list = [];
      for (let d = 0; d < D; d++) list.push(new Acc());
      accs.set(key, list);
    }
    for (let d = 0; d < D; d++) {
      const fld = table.dataFields[d].fld;
      const raw = record[fld];
      const value = typeof raw === 'number' ? (fields[fld].items as PivotValue[])[raw] : raw;
      list[d].add(value || { t: 'm' });
    }
  };
  // Each record adds to a total of every level of its rows with every level of its columns.
  if (after.records.length * (n + 1) * (m + 1) * D > PIVOT_LIMITS.work) refuse(table, 'it would take too long to add up (too many records for its fields and values)');
  const rowPath: number[] = new Array(n);
  const colPath: number[] = new Array(m);
  for (const record of after.records) {
    let shown = true;
    for (let k = 0; k < n && shown; k++) {
      const i = itemOf(record, rows[k]);
      const list = items[rows[k]];
      if (i < 0 || !list || list[i].h) shown = false;
      rowPath[k] = i;
    }
    for (let k = 0; k < m && shown; k++) {
      const i = itemOf(record, cols[k]);
      const list = items[cols[k]];
      if (i < 0 || !list || list[i].h) shown = false;
      colPath[k] = i;
    }
    for (const page of pages) {
      if (!shown) break;
      const i = itemOf(record, page.f);
      const list = items[page.f];
      if (i < 0 || !list) shown = false;
      else if (page.only !== undefined ? i !== page.only : list[i].h) shown = false;
    }
    if (!shown) continue;
    let rowKey = '';
    for (let k = 0; k <= n; k++) {
      if (k > 0) {
        const parent = rowKey;
        let set = children.get(`r${parent}`);
        if (!set) children.set(`r${parent}`, (set = new Set()));
        set.add(rowPath[k - 1]);
        rowKey = k === 1 ? `${rowPath[0]}` : `${rowKey},${rowPath[k - 1]}`;
      }
      let colKey = '';
      for (let e = 0; e <= m; e++) {
        if (e > 0) {
          if (k === 0) {
            let set = children.get(`c${colKey}`);
            if (!set) children.set(`c${colKey}`, (set = new Set()));
            set.add(colPath[e - 1]);
          }
          colKey = e === 1 ? `${colPath[0]}` : `${colKey},${colPath[e - 1]}`;
        }
        add(`${rowKey}|${colKey}`, record);
      }
    }
  }
  /** The items under a line, in the order of the field: those with data, or (showAll) every one that is not hidden. */
  const under = (direction: 'r' | 'c', path: number[], f: number): number[] => {
    const list = items[f] as PivotItem[];
    const out: number[] = [];
    if (table.fields[f].showAll) {
      for (let i = 0; i < list.length; i++) if (list[i].t === undefined && !list[i].h) out.push(i);
      return out;
    }
    const set = children.get(direction + path.join(','));
    if (!set) return out;
    for (let i = 0; i < list.length; i++) if (set.has(i)) out.push(i);
    return out;
  };
  let budget = MAX_PIVOT_OUTPUT_CELLS;
  const spend = () => {
    if (--budget < 0) refuse(table, `it would have more than ${MAX_PIVOT_OUTPUT_CELLS} cells`);
  };
  const rowLines: Line[] = [];
  const walkRows = (level: number, prefix: number[]): void => {
    const f = rows[level];
    const field = table.fields[f];
    for (const item of under('r', prefix, f)) {
      const path = prefix.concat(item);
      spend();
      if (level === n - 1) {
        rowLines.push({ kind: 'item', path, d: 0 });
        continue;
      }
      if (form !== 'tabular') rowLines.push({ kind: 'item', path, d: 0 });
      walkRows(level + 1, path);
      if (field.defaultSubtotal && (form === 'tabular' || !field.subtotalTop)) rowLines.push({ kind: 'sub', path, d: 0 });
    }
  };
  walkRows(0, []);
  // The grand totals of the columns stand in a row at the bottom (colGrandTotals); those of the rows in a column at the right (rowGrandTotals).
  if (table.colGrandTotals) rowLines.push({ kind: 'grand', path: [], d: 0 });
  const colLines: Line[] = [];
  if (m === 0) {
    if (D === 1) colLines.push({ kind: 'data', path: [], d: 0 });
    else for (let d = 0; d < D; d++) colLines.push({ kind: 'value', path: [], d });
  } else {
    const walkCols = (level: number, prefix: number[]): void => {
      const f = cols[level];
      const field = table.fields[f];
      for (const item of under('c', prefix, f)) {
        const path = prefix.concat(item);
        spend();
        if (level === m - 1) {
          colLines.push({ kind: 'item', path, d: 0 });
          continue;
        }
        walkCols(level + 1, path);
        if (field.defaultSubtotal) colLines.push({ kind: 'sub', path, d: 0 });
      }
    };
    walkCols(0, []);
    if (table.rowGrandTotals) colLines.push({ kind: 'grand', path: [], d: 0 });
  }
  const height = shape.firstDataRow + rowLines.length;
  const width = shape.labelCols + colLines.length;
  if (height * width > MAX_PIVOT_OUTPUT_CELLS) refuse(table, `it would have more than ${MAX_PIVOT_OUTPUT_CELLS} cells`);
  const top = old.location.ref.top;
  const left = old.location.ref.left;
  if (top + height > MAX_ROWS || left + width > MAX_COLS) refuse(table, 'it would reach past the end of the sheet');

  const cells: { row: number; col: number; cell: Cell }[] = [];
  const rowKeys = rowLines.map((line) => (line.kind === 'grand' ? '' : line.path.join(',')));
  const colKeys = colLines.map((line) => (line.kind === 'item' || line.kind === 'sub' ? line.path.join(',') : ''));
  const label = (f: number, i: number, role: string): Cell => {
    const item = (items[f] as PivotItem[])[i];
    const named = itemCaption(item);
    const style = styleFor(harvested, role, undefined);
    if (named !== undefined) return { kind: 'string', value: named, style };
    const value = item.x !== undefined && fields[f].items ? (fields[f].items as PivotValue[])[item.x] : undefined;
    if (!value || value.t === 'm') return { kind: 'string', value: harvested.captions.blank ?? WORDS.blank, style };
    switch (value.t) {
      case 'n':
        return { kind: 'number', value: value.v, style: labelStyle(f, style, false) };
      case 'd':
        return { kind: 'number', value: value.v, style: labelStyle(f, style, true) };
      case 'b':
        return { kind: 'boolean', value: value.v, style };
      case 'e':
        return { kind: 'error', value: value.v, style };
      default:
        return { kind: 'string', value: value.v, style };
    }
  };
  /** The style of the label of a number or a date: with the number format of its field, when the field has one. */
  const labelStyle = (f: number, style: number, date: boolean): number => {
    const numFmtId = table.fields[f].numFmtId;
    if (numFmtId === undefined) return date ? dateStyle(workbook, style) : style;
    const format = workbook.styles.cellXfs[style] || workbook.styles.cellXfs[0];
    if (format && format.numFmtId === numFmtId) return style;
    const found = format ? workbook.styles.cellXfs.findIndex((other) => sameFormat(other, { ...format, numFmtId })) : -1;
    if (found < 0) refuse(table, `the labels of the field "${fields[f] ? fields[f].name : f}" have a number format the style sheet has no style with`);
    return found;
  };
  const totalOf = (f: number, i: number, role: string): Cell => {
    const item = (items[f] as PivotItem[])[i];
    const value = item.x !== undefined && fields[f].items ? (fields[f].items as PivotValue[])[item.x] : undefined;
    const text = itemCaption(item) ?? (!value || value.t === 'm' ? harvested.captions.blank ?? WORDS.blank : labelText(workbook, value));
    const words = harvested.total || subtotalWords(table.captions.grandTotal ?? harvested.captions.grand);
    return { kind: 'string', value: words.before + text + words.after, style: styleFor(harvested, role, undefined) };
  };
  place(table, shape, rowLines, colLines, (at) => {
    let cell: Cell | undefined;
    if (at.caption === 'rowHeader') cell = { kind: 'string', value: table.captions.rowHeader ?? harvested.captions.rowHeader ?? WORDS.rowHeader, style: styleFor(harvested, at.role, undefined) };
    else if (at.caption === 'colHeader') cell = { kind: 'string', value: table.captions.colHeader ?? harvested.captions.colHeader ?? WORDS.colHeader, style: styleFor(harvested, at.role, undefined) };
    else if (at.caption === 'grand') cell = { kind: 'string', value: table.captions.grandTotal ?? harvested.captions.grand ?? WORDS.grand, style: styleFor(harvested, at.role, undefined) };
    else if (at.caption === 'data') cell = { kind: 'string', value: table.dataFields[at.d || 0].name, style: styleFor(harvested, at.role, undefined) };
    else if (at.caption === 'field') {
      const f = at.field as number;
      cell = { kind: 'string', value: table.fields[f].name ?? fields[f].name, style: styleFor(harvested, at.role, undefined) };
    } else if (at.rowLine !== undefined && at.colLine !== undefined) {
      const d = at.d || 0;
      const accsHere = accs.get(`${rowKeys[at.rowLine]}|${colKeys[at.colLine]}`);
      const result = accsHere ? accsHere[d].result(table.dataFields[d].subtotal) : undefined;
      const own = [`${at.role}:${d}`, `value:${d}`, `subValue:${d}`, `grandValue:${d}`].some((key) => harvested.styles.has(key));
      const style = valueStyle(workbook, styleFor(harvested, at.role, d), table.dataFields[d].numFmtId, own);
      // A cell without a value shows the caption of missing values (missingCaption), one with an error that of errors (errorCaption).
      if (result === undefined) {
        if (table.missingCaption === undefined) return;
        cell = { kind: 'string', value: table.missingCaption, style };
      } else if (typeof result === 'number') cell = { kind: 'number', value: result, style };
      else cell = table.errorCaption !== undefined ? { kind: 'string', value: table.errorCaption, style } : { kind: 'error', value: result, style };
    } else if (at.field !== undefined && at.item !== undefined) {
      cell = /Sub/.test(at.role) ? totalOf(at.field, at.item, at.role) : label(at.field, at.item, at.role);
    }
    if (cell) cells.push({ row: at.row, col: at.col, cell });
  });

  const location: PivotLocation = {
    ref: { top, left, bottom: top + height - 1, right: left + width - 1 },
    firstHeaderRow: shape.firstHeaderRow,
    firstDataRow: shape.firstDataRow,
    firstDataCol: shape.labelCols,
  };
  return { table, sheet: sheetIndex, state: { location, items, rowItems: fileLines(rowLines), colItems: fileLines(colLines), pageItems }, cells };
}

// ---------------------------------------------------------------------------------------------
// Commands

const intersects = (a: CellRange, b: CellRange) => a.top <= b.bottom && a.bottom >= b.top && a.left <= b.right && a.right >= b.left;
const inside = (range: CellRange, row: number, col: number) => row >= range.top && row <= range.bottom && col >= range.left && col <= range.right;

/** Whether a cell holds anything a table would overwrite. */
const holds = (cell: Cell | undefined): boolean => !!cell && (cell.value !== null || cell.formula !== undefined);

const sameCell = (a: Cell | undefined, b: Cell): boolean => !!a && a.kind === b.kind && a.value === b.value && a.style === b.style && a.formula === undefined && a.runs === undefined;

/** The cache a table is made from and every table made from it. */
function family(workbook: Workbook, index: number, look: boolean): { cache: PivotCache; cacheIndex: number; tables: number[] } {
  const pivots = workbook.pivots;
  const table = pivots ? pivots.tables[index] : undefined;
  if (!pivots || !table || !(index >= 0) || Math.floor(index) !== index) throw new RangeError(`No pivot table at index ${index}`);
  // What else refers to the caches, as the workbook is now (see lookThroughPivots()).
  if (look) lookThroughPivots(workbook);
  const cache = table.cache;
  const cacheIndex = pivots.caches.indexOf(cache);
  if (cacheIndex < 0) refuse(table, 'its cache is not known');
  const tables: number[] = [];
  pivots.tables.forEach((other, i) => {
    if (other.cache === cache) tables.push(i);
  });
  return { cache, cacheIndex, tables };
}

/**
 * The operations that refresh a pivot table: its cache is read again from the source, and every
 * table made from that cache is built again. Throws PivotRefreshError (a RangeError) with the
 * reason when that is not done here, when a table would grow over cells that hold something,
 * and when the source does not fit the cache (another header, another number of columns).
 * Throws RangeError for an index that is no table.
 */
export function refreshPivot(workbook: Workbook, index: number): Op[] {
  return refreshOps(workbook, index, true);
}

/** refreshPivot(), without looking through the workbook when `look` is false (refreshAllPivots() looked through it already). */
function refreshOps(workbook: Workbook, index: number, look: boolean): Op[] {
  const { cache, cacheIndex, tables } = family(workbook, index, look);
  const pivots = workbook.pivots as NonNullable<Workbook['pivots']>;
  const first = pivots.tables[index];
  if (cache.refusal !== undefined) refuse(first, cache.refusal);
  if (cache.referred !== undefined) refuse(first, cache.referred);
  for (const i of tables) {
    const table = pivots.tables[i];
    if (table.refusal !== undefined) refuse(table, table.refusal);
    const sheetIndex = workbook.sheets.indexOf(table.sheet);
    if (sheetIndex < 0) refuse(table, 'its sheet is not in the workbook');
    if (table.sheet.kind !== 'worksheet') refuse(table, 'it does not stand on a worksheet');
  }
  const rows = readSource(workbook, cache, first);
  const before = cacheFieldsNow(cache);
  const after = buildCache(cache, rows, tables.map((i) => pivots.tables[i]));
  const built = tables.map((i) => build(workbook, pivots.tables[i], workbook.sheets.indexOf(pivots.tables[i].sheet), after, before));

  const ops: Op[] = [];
  for (const one of built) {
    const sheet = one.table.sheet;
    const was = tableState(one.table).location.ref;
    const now = one.state.location.ref;
    // Other tables of the sheet, as they will stand.
    for (const other of pivots.tables) {
      if (other === one.table || other.sheet !== sheet) continue;
      const mate = built.find((b) => b.table === other);
      const area = mate ? mate.state.location.ref : tableState(other).location.ref;
      if (intersects(area, now)) refuse(one.table, `it would grow over the pivot table "${other.name}"`);
    }
    for (const merge of sheet.merges) if (intersects(merge, now)) refuse(one.table, `it would grow over the merged cells ${rangeName(merge)}`);
    // Cells beyond the table as it was must be empty, as in Excel.
    for (let r = now.top; r <= now.bottom; r++) {
      const row = sheet.rows.get(r);
      if (!row) continue;
      for (let c = now.left; c <= now.right; c++) {
        if (inside(was, r, c)) continue;
        if (holds(row.cells[c])) refuse(one.table, `it would grow over ${cellName(r, c)}, which is not empty`);
      }
    }
    const wanted = new Map<number, Cell>();
    for (const placed of one.cells) wanted.set((now.top + placed.row) * MAX_COLS + now.left + placed.col, placed.cell);
    const changes: CellChange[] = [];
    const visit = (r: number, c: number) => {
      const key = r * MAX_COLS + c;
      const next = wanted.get(key);
      const current = getCell(sheet, r, c);
      if (next === undefined) {
        if (current !== undefined) changes.push({ row: r, col: c, cell: undefined });
      } else if (!sameCell(current, next)) changes.push({ row: r, col: c, cell: next });
    };
    for (let r = Math.min(was.top, now.top); r <= Math.max(was.bottom, now.bottom); r++) {
      for (let c = Math.min(was.left, now.left); c <= Math.max(was.right, now.right); c++) {
        if (inside(was, r, c) || inside(now, r, c)) visit(r, c);
      }
    }
    if (changes.length > 0) ops.push({ type: 'setCells', sheet: one.sheet, cells: changes });
  }
  // Nothing changes: no cell, and the cache and its tables are what the last refresh made of them.
  if (ops.length === 0 && cache.now && sameJson(cache.now, after) && built.every((one) => one.table.now !== undefined && sameJson(one.table.now, one.state))) return [];
  // The state first: what cannot take it (a workbook edited together) refuses the step before any cell is changed.
  ops.unshift({
    type: 'setPivot',
    sheet: workbook.sheets.indexOf(first.sheet),
    cache: cacheIndex,
    state: after,
    tables: built.map((one) => ({ table: pivots.tables.indexOf(one.table), state: one.state })),
  });
  return ops;
}

/** Whether two states hold the same: of plain data (numbers, text, arrays, objects without cycles). */
function sameJson(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (typeof a !== 'object' || typeof b !== 'object' || a === null || b === null) return typeof a === 'number' && typeof b === 'number' && a !== a && b !== b;
  if (Array.isArray(a) !== Array.isArray(b)) return false;
  if (Array.isArray(a)) {
    const other = b as unknown[];
    if (a.length !== other.length) return false;
    for (let i = 0; i < a.length; i++) if (!sameJson(a[i], other[i])) return false;
    return true;
  }
  const left = a as Record<string, unknown>;
  const right = b as Record<string, unknown>;
  const keys = Object.keys(left).filter((key) => left[key] !== undefined);
  if (keys.length !== Object.keys(right).filter((key) => right[key] !== undefined).length) return false;
  for (const key of keys) if (!sameJson(left[key], right[key])) return false;
  return true;
}

/** Whether a cache reads cells that hold formulas: their values may wait for what another refresh writes. */
function sourceHasFormulas(workbook: Workbook, cache: PivotCache, table: PivotTable): boolean {
  const { sheet, range } = sourceRange(workbook, cache, table);
  for (let r = range.top; r <= range.bottom; r++) {
    const row = sheet.rows.get(r);
    if (!row) continue;
    for (let c = range.left; c <= range.right; c++) if (row.cells[c] && row.cells[c].formula) return true;
  }
  return false;
}

/** Why a pivot table cannot be refreshed now; undefined when it can. Reads the source to tell. */
export function refreshPivotProblem(workbook: Workbook, index: number): string | undefined {
  try {
    refreshPivot(workbook, index);
    return undefined;
  } catch (error) {
    if (error instanceof RangeError) return error.message;
    throw error;
  }
}

/** What refreshing every pivot table of a workbook comes to. */
export interface RefreshAll {
  /** The operations of every refresh that can be made, one cache after the other. */
  ops: Op[];
  /** The tables that are refreshed by them. */
  refreshed: number[];
  /** The tables that are not, each with the reason. */
  refused: { index: number; name: string; reason: string }[];
}

/**
 * Refreshes every pivot table that can be: one cache after another, each with the tables made
 * from it. A cache that cannot be refreshed leaves its tables as they are, and says why; the
 * others are refreshed all the same. The operations of a cache are built on the workbook as the
 * caches before it leave it (a table may read what another one shows): they are applied here
 * while the next is built, and taken back before this returns, so that the caller applies all
 * of them as one step. The workbook is then as it was, its count of edits included.
 */
export function refreshAllPivots(workbook: Workbook): RefreshAll {
  const pivots = workbook.pivots;
  const out: RefreshAll = { ops: [], refreshed: [], refused: [] };
  if (!pivots) return out;
  // One look for every cache: what is applied between two of them (cells, and the state of caches and tables) changes
  // neither the charts of the model nor the package, which is all that is looked at.
  lookThroughPivots(workbook);
  const edits = editsOf(workbook);
  const caches: PivotCache[] = [];
  for (const table of pivots.tables) if (caches.indexOf(table.cache) < 0) caches.push(table.cache);
  const undo: Op[] = [];
  try {
    caches.forEach((cache, k) => {
      const members: number[] = [];
      pivots.tables.forEach((table, j) => {
        if (table.cache === cache) members.push(j);
      });
      let ops: Op[];
      try {
        // Formulas of the source are not calculated again here, between two refreshes: one that may read what an earlier refresh
        // of this step wrote waits for a refresh of its own (whoever refreshes one table calculates the workbook first).
        if (out.ops.some((op) => op.type === 'setCells') && sourceHasFormulas(workbook, cache, pivots.tables[members[0]])) {
          refuse(pivots.tables[members[0]], 'its source has formulas, and another pivot table refreshed with it changed cells: refresh it on its own');
        }
        ops = refreshOps(workbook, members[0], false);
      } catch (error) {
        if (!(error instanceof RangeError)) throw error;
        for (const j of members) out.refused.push({ index: j, name: pivots.tables[j].name, reason: error.message });
        return;
      }
      out.ops.push(...ops);
      out.refreshed.push(...members);
      // The last one is not applied: nothing is built after it.
      if (k < caches.length - 1) for (const op of ops) undo.push(applyOp(workbook, op));
    });
  } finally {
    for (let i = undo.length - 1; i >= 0; i--) applyOp(workbook, undo[i]);
    restoreEdits(workbook, edits);
  }
  return out;
}
