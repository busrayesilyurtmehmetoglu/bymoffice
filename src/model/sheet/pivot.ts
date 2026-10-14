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

import type { CellRange, Sheet, Workbook } from './types';

/** The functions a value of a pivot table can be summarized by, as the file names them. */
export const PIVOT_AGGREGATES = ['sum', 'count', 'average', 'max', 'min', 'product', 'countNums', 'stdDev', 'stdDevp', 'var', 'varp'] as const;
export type PivotAggregate = (typeof PIVOT_AGGREGATES)[number];

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

/** What a refresh made of a table. */
export interface PivotTableState {
  location: PivotLocation;
  /** The items of each field, in the order of PivotTable.fields; undefined for one that lists none. */
  items: (PivotItem[] | undefined)[];
  rowItems: PivotLine[];
  colItems: PivotLine[];
  /** The item each page field shows, in the order of PivotTable.pageFields. */
  pageItems: (number | undefined)[];
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
}

/** The pivot tables of a workbook and their caches. Neither list changes after the file is read. */
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
}

/** The state of a table: what the last refresh made of it, else what the file has. */
export function tableState(table: PivotTable): PivotTableState {
  if (table.now) return table.now;
  return {
    location: table.location,
    items: table.fields.map((field) => field.items),
    rowItems: table.rowItems,
    colItems: table.colItems,
    pageItems: table.pageFields.map((page) => page.item),
  };
}

/** The fields of a cache: as the last refresh made them, else as the file has them. */
export function cacheFieldsNow(cache: PivotCache): PivotCacheField[] {
  return cache.now ? cache.now.fields : cache.fields;
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
