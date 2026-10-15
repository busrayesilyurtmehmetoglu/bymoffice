// Pivot tables of a workbook, as far as the engine knows them.
//
// What a pivot table shows is in the cells of its sheet: the file has it there, and a file that
// is opened shows those cells as they are. The model below is what it takes to build them again
// from the source (a refresh, see src/ops/pivots.ts): the cache the table is made from (its
// fields and the items of each), and the table itself (which fields stand in rows, columns,
// pages and values, which items are hidden, how totals are shown, where it stands).
//
// The parts themselves are written back as the file has them until the table is refreshed;
// then the writer (src/formats/xlsx/pivot-write.ts) changes what a refresh changes and keeps
// every other element and attribute of the parts. A refresh is not applied by changing these
// objects: each holds what the file says, and `now` holds what a refresh made of it (undone,
// `now` goes back to what it was; see the "setPivot" operation).

import { MAX_COLS, MAX_ROWS, foldSheetName, parseRange, type CellRange, type Sheet, type Workbook } from './types';

/** The functions a value of a pivot table can be summarized by, as the file names them. */
export const PIVOT_AGGREGATES = ['sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp'] as const;
export type PivotAggregate = (typeof PIVOT_AGGREGATES)[number];

/** The words Excel uses for the function a value is summarized by, in its default names ("Sum of Sales"). */
const AGGREGATE_WORDS: Record<PivotAggregate, string> = {
  sum: 'Sum',
  count: 'Count',
  average: 'Average',
  max: 'Max',
  min: 'Min',
  product: 'Product',
  countNums: 'Count',
  stdDev: 'StdDev',
  stdDevp: 'StdDevp',
  var: 'Var',
  varp: 'Varp',
};

/** The name Excel gives a value of a field summarized by a function: "Sum of Sales". */
export function defaultValueName(subtotal: PivotAggregate, fieldName: string): string {
  return `${AGGREGATE_WORDS[subtotal]} of ${fieldName}`;
}

/** Pivot tables one workbook holds at most; the rest are kept as the file has them, and are not refreshed. */
export const MAX_PIVOT_TABLES = 1000;
/** Fields of one cache at most (Excel's limit on columns). */
export const MAX_PIVOT_FIELDS = 16384;
/** Items the caches and tables of one workbook list at most, all of them together (shared items, items of fields, lines of rows and columns). */
export const MAX_PIVOT_ITEMS = 4000000;
/**
 * Bytes one part of a pivot table or of its cache definition may have to be read into the model.
 * The part is read as a tree, which takes many times its size; a cache of a million different
 * values of text fits.
 */
export const MAX_PIVOT_PART_BYTES = 16 * 1024 * 1024;
/** Cells of the source a refresh reads at most (rows below the header times columns). */
export const MAX_PIVOT_SOURCE_CELLS = 4000000;
/** Additions to totals a refresh makes at most for one table (records times the levels of rows and columns times the values). */
export const MAX_PIVOT_WORK = 100000000;
/**
 * What looking through the package for what refers to a cache reads at most (see the reader):
 * relationship parts, bytes of one relationships part and of one chart part, bytes of all
 * relationships and of all charts. A test may lower them to reach them with small files.
 */
export const PIVOT_SCAN_LIMITS = {
  relsParts: 10000,
  relsPartBytes: 4 * 1024 * 1024,
  chartPartBytes: 8 * 1024 * 1024,
  relsBytes: 64 * 1024 * 1024,
  chartBytes: 64 * 1024 * 1024,
  /** Relationships and charts together. */
  totalBytes: 64 * 1024 * 1024,
  /** Elements of one chart the reader looks through for its pivot source. */
  chartNodes: 1000000,
};
/** Cells a refresh writes and clears at most, for one table. */
export const MAX_PIVOT_OUTPUT_CELLS = 1000000;

/**
 * One value of a cache: text, a number, a date (as its serial number), a logical value, an
 * error (its text) or nothing (a blank cell of the source).
 */
export type PivotValue = { t: 's'; v: string } | { t: 'n'; v: number } | { t: 'd'; v: number } | { t: 'b'; v: boolean } | { t: 'e'; v: string } | { t: 'm' };

/** A field of a cache: a column of the source. */
export interface PivotCacheField {
  name: string;
  /**
   * The values of the field each once, as the file lists them (<sharedItems>): the items of the
   * field, which records and the items of a table name by their place here. Undefined for a field
   * whose values the records hold themselves (a field of numbers that stands in no row or column).
   */
  items?: PivotValue[];
  /** The number format the file gives the field (cacheField numFmtId: that of its source, for a field of dates), when it is not the general one. */
  numFmtId?: number;
}

/** Where the cache is read from: a range of a sheet of the workbook, or a defined name that stands for one. */
export type PivotSource = { kind: 'range'; sheet: string; ref: string } | { kind: 'name'; name: string; sheet?: string };

/** What a refresh made of a cache. */
export interface PivotCacheState {
  /** The fields as the source has them now, in the order of the cache. */
  fields: PivotCacheField[];
  /**
   * The records, one per row of the source below its header: for each field the place of its
   * value among the items of the field (a field that lists them), or the value itself.
   */
  records: (number | PivotValue)[][];
}

/** A cache of pivot tables (pivotCacheDefinition): what one or more tables are made from. */
export interface PivotCache {
  cacheId: number;
  /** The part of the cache definition, and the part of its records (none when the file saves no data). */
  part: string;
  recordsPart?: string;
  source?: PivotSource;
  /** The fields as the file has them. */
  fields: PivotCacheField[];
  /** Why the cache cannot be refreshed here; undefined when it can. */
  refusal?: string;
  /** What refers to the cache by the places of its items, as the workbook was last looked through (WorkbookPivots.lookThrough); undefined: nothing, or not looked through. */
  referred?: string;
  /** What the last refresh made of it; undefined while it is as the file has it. */
  now?: PivotCacheState;
  /**
   * The range of its source (a source of kind "range") after rows and columns of its sheet were
   * inserted or deleted, as the file writes it; undefined while it is the one the file has. A
   * refresh reads this range, and the writer puts it in the place of the file's (see
   * src/ops/pivot-places.ts). The records are not changed with it: Excel does not either.
   */
  sourceRef?: string;
  /** Set by the reader: why the places its source names cannot follow rows and columns that move; undefined when they can. */
  placesFixed?: string;
  /** For a cache made here and in no file: its definition and its records as they were made (see PivotTable.created). */
  created?: { definition: string; records: string };
}

/** An item of a field of a pivot table (<item>). */
export interface PivotItem {
  /** The place of its value among the items of the cache field; undefined for an item that stands for no value (a subtotal). */
  x?: number;
  /** Its kind: undefined for an item of a value ("data"), "default" for the subtotal of the field, and so on. */
  t?: string;
  /** Hidden: the item is filtered out. */
  h?: boolean;
  /** The attributes of the element as the file has them, in its order: written back as they are, with `x` brought in line. */
  attrs: [string, string][];
}

export type PivotAxis = 'axisRow' | 'axisCol' | 'axisPage' | 'axisValues';

/** A field of a pivot table (<pivotField>): one per field of its cache, in the same order. */
export interface PivotField {
  axis?: PivotAxis;
  /** The caption the table shows for the field, when the file gives one. */
  name?: string;
  /** Undefined for a field that lists no items. */
  items?: PivotItem[];
  compact: boolean;
  outline: boolean;
  subtotalTop: boolean;
  defaultSubtotal: boolean;
  /** Items without data are shown as well. */
  showAll: boolean;
  sortType: 'manual' | 'ascending' | 'descending';
  /** Items new to the source come in hidden (hideNewItems). */
  hideNewItems?: boolean;
  /** The number format the labels of its numbers and dates are shown with (numFmtId), when it is not the general one. */
  numFmtId?: number;
}

/** A value of a pivot table (<dataField>). */
export interface PivotDataField {
  name: string;
  /** The field it summarizes. */
  fld: number;
  subtotal: PivotAggregate;
  numFmtId?: number;
}

/** A field that filters the whole table from above it (<pageField>). */
export interface PivotPageField {
  fld: number;
  /** The one item shown (its place among the items of the field); undefined: all items that are not hidden. */
  item?: number;
}

/** One line of the rows or the columns of a table (<i> of <rowItems> or <colItems>). */
export interface PivotLine {
  /** undefined: a line of items ("data"); "default": a subtotal; "grand": the grand total; others as the file names them. */
  t?: string;
  /** How many fields at the start of the line are those of the line before. */
  r: number;
  /** The value (data field) the line is of, where the values stand in this direction. */
  i: number;
  /** The places of its items among the items of each field, from the field after the `r` repeated ones on. */
  x: number[];
}

/** Where a table stands (<location>), in the cells of its sheet. */
export interface PivotLocation {
  /** The body of the table: its header, its labels and its values; the fields of the page stand above it. */
  ref: CellRange;
  firstHeaderRow: number;
  firstDataRow: number;
  firstDataCol: number;
}

/**
 * Which fields of a table stand where (rows, columns, page, values) and how they are totalled: what
 * changing the fields of a table sets (src/ops/pivot-fields.ts). The table is built again from it.
 */
export interface PivotTableDef {
  /** One per field of the cache, as PivotTable.fields; their items are those of the state (PivotTableState.items). */
  fields: PivotField[];
  rowFields: number[];
  colFields: number[];
  pageFields: PivotPageField[];
  dataFields: PivotDataField[];
  rowGrandTotals: boolean;
  colGrandTotals: boolean;
  /** As PivotTable.pages: the rows and columns the fields of the page take above the body. */
  pages?: { rows: number; cols: number };
}

/** What a refresh made of a table. */
export interface PivotTableState {
  location: PivotLocation;
  /** The items of each field, in the order of PivotTable.fields; undefined for one that lists none. */
  items: (PivotItem[] | undefined)[];
  rowItems: PivotLine[];
  colItems: PivotLine[];
  /** The item each page field shows, in the order of the page fields (those of `def` when it is set). */
  pageItems: (number | undefined)[];
  /** The fields as they were changed (see PivotTableDef); undefined while they are those of the file (or, for a table made here, those it was made with). */
  def?: PivotTableDef;
}

/** Captions the file gives the table in place of the words of the application. */
export interface PivotCaptions {
  rowHeader?: string;
  colHeader?: string;
  grandTotal?: string;
  data?: string;
}

/** A pivot table (pivotTableDefinition). */
export interface PivotTable {
  name: string;
  part: string;
  /** The sheet the table stands on. */
  sheet: Sheet;
  cache: PivotCache;
  location: PivotLocation;
  fields: PivotField[];
  /** Fields of the rows and of the columns, in order; -2 stands for the values. */
  rowFields: number[];
  colFields: number[];
  pageFields: PivotPageField[];
  dataFields: PivotDataField[];
  rowItems: PivotLine[];
  colItems: PivotLine[];
  rowGrandTotals: boolean;
  colGrandTotals: boolean;
  captions: PivotCaptions;
  /** What a cell of values without a value shows (missingCaption, while showMissing); undefined: nothing. */
  missingCaption?: string;
  /** What a cell of values whose value is an error shows (errorCaption, while showError); undefined: the error. */
  errorCaption?: string;
  /** Items are sorted by the lists of the application first (months, days) when a field is sorted (customListSort). */
  customListSort: boolean;
  /** Why the table cannot be refreshed here; undefined when it can. */
  refusal?: string;
  /** What the last refresh made of it; undefined while it is as the file has it. */
  now?: PivotTableState;
  /**
   * The body of the table after rows or columns were inserted or deleted above it or to its left,
   * while it is not refreshed (a refresh has its place in `now`, and moves it there as well);
   * undefined while it stands where the file has it.
   */
  placed?: CellRange;
  /** The rows and the columns of fields of the page that stand above the body (location rowPageCount, colPageCount); undefined without such fields. */
  pages?: { rows: number; cols: number };
  /** Set by the reader: why the fields of the table cannot be changed here although it can be refreshed; undefined when they can. */
  editRefusal?: string;
  /**
   * For a table made here (src/ops/pivot-create.ts) and in no file: its definition as it was made,
   * which the writer starts from in the place of the bytes of a file. Its part (`part`) is ''
   * until a save gives it one, as for a sheet that is in no file.
   */
  created?: string;
}

/**
 * The pivot tables of a workbook and their caches. A table made here is added at the end of both
 * lists, and taken out of them again by its undo ("addPivot"): the places of the others never change.
 */
export interface WorkbookPivots {
  caches: PivotCache[];
  tables: PivotTable[];
  /**
   * Set by the reader: looks through the workbook as it is now (the charts of the model, the parts
   * of the package) for what refers to the caches (a table the model could not take in, a slicer, a
   * pivot chart, ...) and says it on each cache it finds referred to (PivotCache.referred). Opening
   * a file does not do it; a refresh does it each time.
   */
  lookThrough?: (workbook: Workbook) => void;
  /**
   * Set by the reader: why rows and columns cannot be inserted or deleted in the workbook because
   * of its pivot tables (a table or a cache of the package that the model does not hold, whose
   * places would not follow); undefined when the model holds every one of them.
   */
  fixed?: string;
}

/** The state of a table: what the last refresh made of it, else what the file has. */
export function tableState(table: PivotTable): PivotTableState {
  if (table.now) return table.now;
  return {
    location: table.placed ? { ...table.location, ref: table.placed } : table.location,
    items: table.fields.map((field) => field.items),
    rowItems: table.rowItems,
    colItems: table.colItems,
    pageItems: table.pageFields.map((page) => page.item),
  };
}

/** The fields of a table and where they stand: as they were last changed, else as the file has them. */
export function tableDef(table: PivotTable): PivotTableDef {
  const def = table.now ? table.now.def : undefined;
  if (def) return def;
  const out: PivotTableDef = {
    fields: table.fields,
    rowFields: table.rowFields,
    colFields: table.colFields,
    pageFields: table.pageFields,
    dataFields: table.dataFields,
    rowGrandTotals: table.rowGrandTotals,
    colGrandTotals: table.colGrandTotals,
  };
  if (table.pages) out.pages = table.pages;
  return out;
}

/** The table with its fields as they stand now (see tableDef()): what a table is built by. The table itself when they are those of the file. */
export function tableNow(table: PivotTable): PivotTable {
  const def = table.now ? table.now.def : undefined;
  if (!def) return table;
  const out: PivotTable = { ...table, ...def };
  if (!def.pages) delete out.pages;
  return out;
}

/** The fields of a cache: as the last refresh made them, else as the file has them. */
export function cacheFieldsNow(cache: PivotCache): PivotCacheField[] {
  return cache.now ? cache.now.fields : cache.fields;
}

/** The source of a cache as it stands now: with the range that rows and columns which moved left it (PivotCache.sourceRef). */
export function cacheSourceNow(cache: PivotCache): PivotSource | undefined {
  const source = cache.source;
  return source && source.kind === 'range' && cache.sourceRef !== undefined ? { ...source, ref: cache.sourceRef } : source;
}

function columnIndex(letters: string): number {
  let n = 0;
  for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.toUpperCase().charCodeAt(i) - 64);
  return n - 1;
}

/**
 * The cells the range of a source names, as the file writes it: "A1:C9", one cell, whole columns
 * ("A:C") or whole rows ("1:9"), with or without "$". Undefined for anything else.
 */
export function sourceArea(ref: string): CellRange | undefined {
  const plain = ref.replace(/\$/g, '');
  let match = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(plain);
  if (match) {
    const a = columnIndex(match[1]);
    const b = columnIndex(match[2]);
    if (a >= MAX_COLS || b >= MAX_COLS) return undefined;
    return { top: 0, bottom: MAX_ROWS - 1, left: Math.min(a, b), right: Math.max(a, b) };
  }
  match = /^(\d{1,7}):(\d{1,7})$/.exec(plain);
  if (match) {
    const a = Number(match[1]) - 1;
    const b = Number(match[2]) - 1;
    if (a < 0 || b < 0 || a >= MAX_ROWS || b >= MAX_ROWS) return undefined;
    return { top: Math.min(a, b), bottom: Math.max(a, b), left: 0, right: MAX_COLS - 1 };
  }
  if (!/^[A-Za-z]{1,3}\d{1,7}(?::[A-Za-z]{1,3}\d{1,7})?$/.test(plain)) return undefined;
  return parseRange(plain);
}

/**
 * The columns of a source (`area` on `sheet`, where it stands now) that are fields of its cache,
 * read from the header row of the source as the cells have it now: the header texts matched to
 * the names of the fields, without regard to case, in their order; the columns as absolute
 * indexes. 'all' when the match is not one and whole (a name that is missing, found twice or
 * given to two fields, or names that do not come in the order of the fields): every column of
 * the source is then taken for a field. A column whose header is empty or names no field is none.
 *
 * Nothing about it is kept: it is read anew whenever it is asked for, from the model alone, so
 * that every participant of a shared workbook (who has the same model) comes to the same answer.
 * The header row is read once, over its cells that stand inside the source.
 */
export function fieldColumnsNow(cache: PivotCache, sheet: Sheet, area: CellRange): number[] | 'all' {
  const fields = cacheFieldsNow(cache);
  if (fields.length === 0) return 'all';
  // Each header text once, with its column; -1 for a text two columns have.
  const headers = new Map<string, number>();
  const row = sheet.rows.get(area.top);
  if (row) {
    row.cells.forEach((cell, c) => {
      if (c < area.left || c > area.right) return;
      const value = cell.value;
      const text = typeof value === 'string' ? value.toLowerCase() : typeof value === 'number' ? String(value) : typeof value === 'boolean' ? (value ? 'true' : 'false') : '';
      if (text !== '') headers.set(text, headers.has(text) ? -1 : c);
    });
  }
  const out: number[] = [];
  for (const field of fields) {
    const name = field.name.toLowerCase();
    if (name === '') return 'all';
    // Two fields of one name find the same column: the second is not after the first, and that refuses it below.
    const at = headers.get(name);
    if (at === undefined || at < 0 || (out.length > 0 && at <= out[out.length - 1])) return 'all';
    out.push(at);
  }
  return out;
}

/** A defined name that stands for one range of a sheet: "Data!$A$1:$C$9", "'My data'!A1:C9". */
const NAMED_RANGE = /^(?:'((?:[^']|'')+)'|([^!'"\s]+))!(\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?)$/;

/**
 * The sheet and the range a source of kind "name" stands for: the defined name local to the sheet
 * the source names, else the one of the workbook. A reason, as text, when there is no such name
 * or it stands for something else than one range of a sheet.
 */
export function namedSourceRange(workbook: Workbook, source: { name: string; sheet?: string }): { sheetName: string; ref: string } | string {
  const wanted = source.name.toLowerCase();
  let sheetIndex = -1;
  if (source.sheet !== undefined) {
    const folded = foldSheetName(source.sheet);
    sheetIndex = workbook.sheets.findIndex((sheet) => foldSheetName(sheet.name) === folded);
  }
  const found =
    workbook.definedNames.find((name) => name.name.toLowerCase() === wanted && name.localSheet !== undefined && name.localSheet === sheetIndex) ||
    workbook.definedNames.find((name) => name.name.toLowerCase() === wanted && name.localSheet === undefined);
  if (!found) return `its source "${source.name}" is a table or a name that is not known here`;
  const match = NAMED_RANGE.exec(found.value.trim());
  if (!match) return `its source "${source.name}" is not one range of cells`;
  return { sheetName: match[1] !== undefined ? match[1].replace(/''/g, "'") : match[2], ref: match[3] };
}

/** Whether two values of a cache are one item: text without regard to case, as Excel groups it. */
export function valueKey(value: PivotValue): string {
  switch (value.t) {
    case 's':
      return `s${value.v.toLowerCase()}`;
    case 'n':
      return `n${value.v}`;
    case 'd':
      return `d${value.v}`;
    case 'b':
      return value.v ? 'b1' : 'b0';
    case 'e':
      return `e${value.v}`;
    default:
      return 'm';
  }
}
