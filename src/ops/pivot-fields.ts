// Changing the fields of a pivot table: which fields stand in its rows, columns, page (filters)
// and values, how a value is summarized and shown, subtotals and grand totals, which items are
// hidden, how the items of a field are sorted.
//
// The table is then built again the way a refresh builds it (src/ops/pivots.ts), from its source:
// Excel lays the table out again from the records of its cache, which are not read here when the
// file is opened, so the source is read instead - a change of the fields refreshes the table as
// well. What a refresh refuses, a change refuses: only a table the refresh can build is changed,
// and only into a layout it builds (one or more fields in rows; values in columns, or fields in
// columns with one value). The new fields are kept in the state of the table (PivotTableState.def):
// undone, the table is as it was; saved, the writer puts them into its definition.
//
// Like the other commands these functions only build operations; one change is one undo step.

import {
  MAX_PIVOT_FIELDS,
  PIVOT_AGGREGATES,
  cacheFieldsNow,
  defaultValueName,
  tableDef,
  tableState,
  type PivotAggregate,
  type PivotDataField,
  type PivotField,
  type PivotItem,
  type PivotTable,
  type PivotTableDef,
} from '../model/sheet/pivot';
import type { Workbook } from '../model/sheet/types';
import { formatCodeOf } from '../model/sheet/color';
import { PivotRefreshError, itemText, lookThroughPivots, refreshOps, type PivotEdit } from './pivots';
import { ensureCellFormat, ensureNumberFormat, markStyles, takeAddedStyles, type Op } from './ops';

/** Fields a table may have in its rows and columns together (the reader refuses more: each is a level of the lines). */
const MAX_AXIS_FIELDS = 64;
/** Values one table may have. */
const MAX_VALUES = 256;
/** The longest name of a value. */
const MAX_VALUE_NAME = 255;

/** The order the items of a field are sorted in. */
export type PivotSort = 'manual' | 'ascending' | 'descending';

/** A value of a table as the fields of the table list it. */
export interface PivotValueInfo {
  field: number;
  subtotal: PivotAggregate;
  name: string;
  /** The number format its cells are shown with; undefined: that of the cells. */
  numFmt?: string;
}

/** A field of a table as the panel of its fields shows it. */
export interface PivotFieldInfo {
  /** Its place among the fields of the table (the columns of its source). */
  index: number;
  name: string;
  /** Its items, for a field that lists them: the text each shows, and whether it is hidden. */
  items?: { text: string; hidden: boolean }[];
  /** Whether it shows a subtotal for each item (in rows and columns). */
  subtotal: boolean;
  sort: PivotSort;
}

/** The fields of a table and where they stand: what pivotFields() gives and setPivotFields() changes. */
export interface PivotFieldsInfo {
  /** Why the fields cannot be changed here; undefined when they can. */
  refusal?: string;
  fields: PivotFieldInfo[];
  rows: number[];
  cols: number[];
  /** The fields of the page, each with the one item it shows (its place among the items of the field), if it shows one. */
  filters: { field: number; item?: number }[];
  values: PivotValueInfo[];
  rowGrandTotals: boolean;
  colGrandTotals: boolean;
}

/** A value as a change sets it; what is left out is the value's own (a value of the field the table had), else the default. */
export interface PivotValueChange {
  field: number;
  subtotal?: PivotAggregate;
  /** Its name; by default Excel's ("Sum of Sales"), which follows the function while it is the default. */
  name?: string;
  /** A number format code ("#,##0.00"); "" or "General": that of the cells. */
  numFmt?: string;
}

/** A change of the fields of a table: what is left out stays as it is. */
export interface PivotFieldsChange {
  rows?: number[];
  cols?: number[];
  filters?: (number | { field: number; item?: number })[];
  values?: PivotValueChange[];
  rowGrandTotals?: boolean;
  colGrandTotals?: boolean;
  /** Subtotals turned on or off, by field. */
  subtotals?: { field: number; on: boolean }[];
  /** The items of a field that are hidden (places among its items, as pivotFields() lists them): all others are shown. */
  hidden?: { field: number; items: number[] }[];
  sort?: { field: number; order: PivotSort }[];
}

/** A change of the fields of a table that is not made, with the reason. */
export class PivotFieldsError extends RangeError {
  constructor(message: string) {
    super(message);
    this.name = 'PivotFieldsError';
  }
}

const refuse = (table: PivotTable, reason: string): never => {
  throw new PivotFieldsError(`Cannot change the fields of the pivot table "${table.name}": ${reason}`);
};

/** Number formats the file names by their number (built into the application), by their code. */
const BUILTIN_CODES: Record<string, number> = { '0': 1, '0.00': 2, '#,##0': 3, '#,##0.00': 4, '0%': 9, '0.00%': 10, '0.00E+00': 11 };

/** The code of the number format a value is shown with; undefined: none of its own. */
function formatCode(workbook: Workbook, numFmtId: number | undefined): string | undefined {
  if (numFmtId === undefined || numFmtId === 0) return undefined;
  return formatCodeOf(workbook.styles, numFmtId);
}

/** The table at `index`, or RangeError. */
function tableAt(workbook: Workbook, index: number): PivotTable {
  const pivots = workbook.pivots;
  const table = pivots && Math.floor(index) === index ? pivots.tables[index] : undefined;
  if (!table) throw new RangeError(`No pivot table at index ${index}`);
  return table;
}

/**
 * Why the fields of a table cannot be changed, without reading its source; undefined when they may
 * be. Looks through the workbook first (see lookThroughPivots()) unless `look` is false: then what
 * the last look found is said.
 */
export function pivotFieldsProblem(workbook: Workbook, index: number, look = true): string | undefined {
  const table = tableAt(workbook, index);
  if (look) lookThroughPivots(workbook);
  const reason = table.cache.refusal ?? table.cache.referred ?? table.refusal ?? table.editRefusal;
  return reason === undefined ? undefined : `Cannot change the fields of the pivot table "${table.name}": ${reason}`;
}

/**
 * The fields of the pivot table at `index` (as listPivots() numbers them) and where they stand. With
 * `look: false` the workbook is not looked through for what refers to the caches (a panel that is
 * shown again after every edit): the refusal is what the last look found.
 */
export function pivotFields(workbook: Workbook, index: number, options: { look?: boolean } = {}): PivotFieldsInfo {
  const table = tableAt(workbook, index);
  const def = tableDef(table);
  const state = tableState(table);
  const cacheFields = cacheFieldsNow(table.cache);
  const out: PivotFieldsInfo = {
    fields: def.fields.map((field, f) => {
      const values = cacheFields[f] ? cacheFields[f].items : undefined;
      const list = state.items[f];
      const info: PivotFieldInfo = { index: f, name: cacheFields[f] ? cacheFields[f].name : '', subtotal: field.defaultSubtotal, sort: field.sortType };
      if (list) {
        info.items = list
          .filter((item) => item.t === undefined)
          .map((item) => ({ text: itemText(workbook, item, values && item.x !== undefined ? values[item.x] : undefined), hidden: !!item.h }));
      }
      return info;
    }),
    rows: def.rowFields.filter((f) => f >= 0),
    cols: def.colFields.filter((f) => f >= 0),
    filters: def.pageFields.map((page, p) => {
      const shown = state.pageItems[p];
      const list = state.items[page.fld];
      // The place among the data items (what `items` above lists), not among all of them.
      const at = shown === undefined || !list ? -1 : list.filter((item) => item.t === undefined).indexOf(list[shown]);
      return at >= 0 ? { field: page.fld, item: at } : { field: page.fld };
    }),
    values: def.dataFields.map((data) => {
      const value: PivotValueInfo = { field: data.fld, subtotal: data.subtotal, name: data.name };
      const code = formatCode(workbook, data.numFmtId);
      if (code !== undefined) value.numFmt = code;
      return value;
    }),
    rowGrandTotals: def.rowGrandTotals,
    colGrandTotals: def.colGrandTotals,
  };
  const reason = pivotFieldsProblem(workbook, index, options.look !== false);
  if (reason !== undefined) out.refusal = reason;
  return out;
}

const isIndex = (value: unknown, count: number): value is number => typeof value === 'number' && value >= 0 && value < count && Math.floor(value) === value;

/**
 * Checks the fields of a layout against `count` fields (whatever the caller gave: everything is
 * checked); returns the reason it is not one that is built here, or undefined.
 */
export function layoutProblem(count: number, rows: unknown, cols: unknown, filters: unknown, values: unknown): string | undefined {
  if (!Array.isArray(rows) || !Array.isArray(cols) || !Array.isArray(filters) || !Array.isArray(values)) return 'the fields are lists';
  if (count > MAX_PIVOT_FIELDS) return 'it has too many fields';
  const seen = new Set<number>();
  for (const f of [...rows, ...cols, ...filters]) {
    if (!isIndex(f, count)) return `there is no field ${String(f).slice(0, 20)}`;
    if (seen.has(f)) return 'a field stands in two places (rows, columns, filters)';
    seen.add(f);
  }
  if (rows.length + cols.length > MAX_AXIS_FIELDS) return `it would have more than ${MAX_AXIS_FIELDS} fields in its rows and columns`;
  if (rows.length === 0) return 'it needs a field in its rows here';
  if (values.length === 0) return 'it needs a value here';
  if (values.length > MAX_VALUES) return `it would have more than ${MAX_VALUES} values`;
  if (values.length > 1 && cols.length > 0) return 'several values and fields in the columns together are not built here';
  return undefined;
}

/** The values of a change as the table states them; the reason as text when one is not valid. */
function valuesOf(workbook: Workbook, def: PivotTableDef, names: string[], change: PivotValueChange[]): PivotDataField[] | string {
  const out: PivotDataField[] = [];
  const used = new Set<number>();
  const taken = new Set<string>();
  for (const one of change) {
    if (!one || typeof one !== 'object' || !isIndex(one.field, names.length)) return 'a value names a field the table does not have';
    if (one.subtotal !== undefined && (PIVOT_AGGREGATES as readonly string[]).indexOf(one.subtotal) < 0) return `a value is summarized by a function that is not known here ("${String(one.subtotal).slice(0, 20)}")`;
    // What is left out is that of a value of the field the table had (the first that is not taken yet).
    const k = def.dataFields.findIndex((data, j) => data.fld === one.field && !used.has(j));
    const had = k >= 0 ? def.dataFields[k] : undefined;
    if (k >= 0) used.add(k);
    const subtotal = one.subtotal ?? (had ? had.subtotal : 'sum');
    const fieldName = names[one.field];
    let name = one.name;
    if (name === undefined) name = had && had.name !== defaultValueName(had.subtotal, fieldName) ? had.name : defaultValueName(subtotal, fieldName);
    if (typeof name !== 'string' || name.trim() === '') return 'a value has no name';
    if (name.length > MAX_VALUE_NAME) return `the name of a value is longer than ${MAX_VALUE_NAME} characters`;
    // Excel takes no name a field has, and no name twice: a name made here gets a number (as Excel's "Sum of Sales2").
    const folded = (text: string) => text.toLowerCase();
    const fieldNamed = (text: string) => names.some((field) => folded(field) === folded(text));
    if (one.name !== undefined && fieldNamed(name)) return `a value cannot have the name of a field ("${name.slice(0, 60)}")`;
    if (one.name !== undefined && taken.has(folded(name))) return `two values have the name "${name.slice(0, 60)}"`;
    if (taken.has(folded(name)) || fieldNamed(name)) {
      const base = name;
      for (let n = 2; taken.has(folded(name)) || fieldNamed(name); n++) name = `${base}${n}`;
    }
    taken.add(folded(name));
    const data: PivotDataField = { name, fld: one.field, subtotal };
    let numFmtId = had ? had.numFmtId : undefined;
    if (one.numFmt !== undefined) {
      if (typeof one.numFmt !== 'string' || one.numFmt.length > 255) return 'a number format is not one';
      const code = one.numFmt.trim();
      numFmtId = code === '' || /^general$/i.test(code) ? undefined : BUILTIN_CODES[code] ?? ensureNumberFormat(workbook.styles, code);
      // A cell of the value finds a style with the format (see valueStyle() of the refresh): it is added to the style sheet, for good, as a format is.
      if (numFmtId !== undefined) ensureCellFormat(workbook.styles, 0, numFmtId);
    }
    if (numFmtId !== undefined) data.numFmtId = numFmtId;
    out.push(data);
  }
  return out;
}

/** The form the fields of the rows of a table are laid out in: what a field put into its rows takes. */
function formOf(def: PivotTableDef): { compact: boolean; outline: boolean } | undefined {
  const first = def.rowFields.find((f) => f >= 0);
  const field = first === undefined ? undefined : def.fields[first];
  return field ? { compact: field.compact, outline: field.outline } : undefined;
}

/**
 * The operations that change the fields of the pivot table at `index` (as listPivots() numbers it)
 * and build it again from its source (with every other table of its cache, refreshed). Throws
 * PivotFieldsError (a RangeError) with the reason when the change is not made here: a table a
 * refresh would not build, a layout that is not built here, a field with too many items for its
 * rows, a table that would grow over cells that hold something. An empty list: nothing changes.
 */
export function setPivotFields(workbook: Workbook, index: number, change: PivotFieldsChange): Op[] {
  // The formats a value asks for are added while the table is built, and taken out again: they come back as the first step.
  const mark = markStyles(workbook.styles);
  let ops: Op[] | undefined;
  try {
    ops = changeFields(workbook, index, change);
  } finally {
    const table = workbook.pivots ? workbook.pivots.tables[index] : undefined;
    const again = takeAddedStyles(workbook.styles, mark, table ? Math.max(0, workbook.sheets.indexOf(table.sheet)) : 0);
    if (ops && ops.length > 0 && again) ops.unshift(again);
  }
  return ops;
}

function changeFields(workbook: Workbook, index: number, change: PivotFieldsChange): Op[] {
  const table = tableAt(workbook, index);
  if (!change || typeof change !== 'object') throw new RangeError('A change of the fields of a pivot table is an object');
  const problem = pivotFieldsProblem(workbook, index);
  if (problem !== undefined) throw new PivotFieldsError(problem);
  const def = tableDef(table);
  const state = tableState(table);
  const names = cacheFieldsNow(table.cache).map((field) => field.name);
  const count = def.fields.length;
  if (names.length !== count) refuse(table, 'its fields are not those of its cache');

  const rows = change.rows !== undefined ? change.rows : def.rowFields.filter((f) => f >= 0);
  const cols = change.cols !== undefined ? change.cols : def.colFields.filter((f) => f >= 0);
  const filters: { field: number; item?: number }[] =
    change.filters !== undefined
      ? (Array.isArray(change.filters) ? change.filters : [change.filters]).map((one) => (typeof one === 'object' && one !== null ? one : { field: one as number }))
      : def.pageFields.map((page) => ({ field: page.fld }));
  const valueChange = change.values !== undefined ? change.values : def.dataFields.map((data) => ({ field: data.fld, subtotal: data.subtotal, name: data.name }));
  const wrong = layoutProblem(count, rows, cols, filters.map((one) => one.field), valueChange);
  if (wrong !== undefined) refuse(table, wrong);
  const dataFields = valuesOf(workbook, def, names, valueChange);
  if (typeof dataFields === 'string') return refuse(table, dataFields);

  // The fields, with where each stands now. A field put into the rows (or the columns) takes the form of those that are there.
  const form = formOf(def);
  const fields: PivotField[] = def.fields.map((field, f) => {
    const next: PivotField = { ...field };
    const axis = rows.indexOf(f) >= 0 ? 'axisRow' : cols.indexOf(f) >= 0 ? 'axisCol' : filters.some((one) => one.field === f) ? 'axisPage' : undefined;
    if (axis) next.axis = axis;
    else delete next.axis;
    const before = def.rowFields.indexOf(f) >= 0 || def.colFields.indexOf(f) >= 0;
    if ((axis === 'axisRow' || axis === 'axisCol') && !before && form) {
      next.compact = form.compact;
      next.outline = form.outline;
    }
    return next;
  });
  for (const one of Array.isArray(change.subtotals) ? change.subtotals : []) {
    if (!one || !isIndex(one.field, count) || typeof one.on !== 'boolean') refuse(table, 'a subtotal names no field');
    fields[one.field] = { ...fields[one.field], defaultSubtotal: one.on };
  }
  for (const one of Array.isArray(change.sort) ? change.sort : []) {
    if (!one || !isIndex(one.field, count) || (one.order !== 'manual' && one.order !== 'ascending' && one.order !== 'descending')) refuse(table, 'a sort names no field or no order');
    fields[one.field] = { ...fields[one.field], sortType: one.order };
  }
  const next: PivotTableDef = {
    fields,
    rowFields: rows.slice(),
    colFields: dataFields.length > 1 ? cols.concat([-2]) : cols.slice(),
    pageFields: filters.map((one) => ({ fld: one.field })),
    dataFields,
    rowGrandTotals: change.rowGrandTotals !== undefined ? change.rowGrandTotals === true : def.rowGrandTotals,
    colGrandTotals: change.colGrandTotals !== undefined ? change.colGrandTotals === true : def.colGrandTotals,
  };
  if (filters.length > 0) next.pages = { rows: filters.length, cols: 1 };

  // The items to start from: hidden or shown as the change says.
  const items: (PivotItem[] | undefined)[] = state.items.slice();
  const dataItems = (f: number): number[] => {
    const list = items[f] || [];
    const out: number[] = [];
    list.forEach((item, i) => {
      if (item.t === undefined) out.push(i);
    });
    return out;
  };
  for (const one of Array.isArray(change.hidden) ? change.hidden : []) {
    if (!one || !isIndex(one.field, count) || !Array.isArray(one.items)) refuse(table, 'hidden items name no field');
    if (!items[one.field]) refuse(table, `the field "${names[one.field].slice(0, 60)}" lists no items yet: put it into the rows, columns or filters first`);
    const places = dataItems(one.field);
    const hide = new Set<number>();
    for (const k of one.items) {
      if (!isIndex(k, places.length)) refuse(table, `the field "${names[one.field].slice(0, 60)}" has no item ${String(k).slice(0, 20)}`);
      hide.add(places[k]);
    }
    if (hide.size === places.length && places.length > 0) refuse(table, `every item of the field "${names[one.field].slice(0, 60)}" would be hidden`);
    items[one.field] = (items[one.field] as PivotItem[]).map((item, i) => {
      if (item.t !== undefined) return item;
      const hidden = hide.has(i);
      if (!!item.h === hidden) return item;
      const attrs = item.attrs.filter(([name]) => name !== 'h');
      const out: PivotItem = { ...item, attrs: hidden ? [['h', '1'] as [string, string]].concat(attrs) : attrs };
      if (hidden) out.h = true;
      else delete out.h;
      return out;
    });
  }
  // What each field of the page shows: the item the change names, else the one it showed.
  const pageItems = filters.map((one) => {
    if (one.item === undefined) {
      if (change.filters !== undefined && Object.prototype.hasOwnProperty.call(one, 'item')) return undefined;
      const p = def.pageFields.findIndex((page) => page.fld === one.field);
      return p >= 0 ? state.pageItems[p] : undefined;
    }
    const places = items[one.field] ? dataItems(one.field) : [];
    if (!isIndex(one.item, places.length)) refuse(table, `the field "${names[one.field].slice(0, 60)}" has no item ${String(one.item).slice(0, 20)} to show`);
    return places[one.item as number];
  });
  // A field of the page shows one item or hides some, not both (what Excel writes): items hidden by this change take
  // the one item away (unless the change names it too), one item shown shows the others again.
  const hiddenNow = new Set((Array.isArray(change.hidden) ? change.hidden : []).map((one) => one.field));
  filters.forEach((one, p) => {
    if (pageItems[p] === undefined) return;
    if (hiddenNow.has(one.field) && one.item === undefined) {
      pageItems[p] = undefined;
      return;
    }
    const list = items[one.field];
    if (!list || !list.some((item) => item.h)) return;
    items[one.field] = list.map((item) => {
      if (!item.h) return item;
      const shown: PivotItem = { ...item, attrs: item.attrs.filter(([name]) => name !== 'h') };
      delete shown.h;
      return shown;
    });
  });

  // The values each takes the style of: one of the table that is the same value.
  const used = new Set<number>();
  const valuesFrom = dataFields.map((data) => {
    const k = def.dataFields.findIndex((old, j) => !used.has(j) && old.fld === data.fld && old.subtotal === data.subtotal && old.numFmtId === data.numFmtId);
    if (k >= 0) used.add(k);
    return k;
  });

  // Room for the fields of the page above the body: a row each and a blank one; the body moves down when there is less.
  const body = state.location.ref;
  const edit: PivotEdit = { def: next, items, pageItems, valuesFrom };
  if (next.pages && body.top < next.pages.rows + 1) edit.at = { top: next.pages.rows + 1, left: body.left };
  try {
    return refreshOps(workbook, index, false, edit);
  } catch (error) {
    if (error instanceof PivotRefreshError) throw new PivotFieldsError(error.message.replace(/^Cannot refresh the pivot table/, 'Cannot change the fields of the pivot table'));
    throw error;
  }
}

/** Why a change of the fields of a table is not made; undefined when it is. Reads the source to tell. */
export function setPivotFieldsProblem(workbook: Workbook, index: number, change: PivotFieldsChange): string | undefined {
  try {
    setPivotFields(workbook, index, change);
    return undefined;
  } catch (error) {
    if (error instanceof RangeError) return error.message;
    throw error;
  }
}
