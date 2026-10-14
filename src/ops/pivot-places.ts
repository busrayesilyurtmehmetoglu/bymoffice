// Where pivot tables and their sources go when rows or columns are inserted or deleted.
//
// A pivot table names cells in two places of the package: the range its cache is read from
// (worksheetSource ref, on the sheet worksheetSource sheet names) and where its body stands on its
// own sheet (location ref; the other attributes of the location count from there). Both follow an
// insertion or a deletion the way Excel moves them:
//
// - The source moves with its cells, and grows by what is inserted inside it (not at its first
//   row or column, nor right after its last). A deletion that takes a column of the source (a field
//   of the cache would name a column that is gone) or its header row (the fields would take their
//   names from another row) is refused, so is one that takes all of it. Which columns are fields is
//   read from the header row of the source as it is at the moment of the edit (fieldColumnsNow() in
//   the model; nothing about it is kept): a column inserted inside the source, whose header is empty
//   or names no field, may be deleted again. A source given by a defined
//   name is left alone: the name follows by itself (the same deletions are refused for it).
// - A table moves with its cells, as a whole, its fields of the page above it included. An edit
//   that would cut through it is refused, as in Excel ("cannot change part of a PivotTable").
//
// The records of a cache are not changed: Excel does not read the source again for an insertion
// either. What moved is kept beside what the file has (PivotCache.sourceRef, PivotTable.placed, and
// the location of a table that was refreshed, PivotTable.now); the writer puts the new ranges in
// the place of the old ones and leaves every other byte of the parts as it is.
//
// Every participant of a shared workbook works this out alike from the edit and the workbook: the
// edit travels alone, nothing of the pivot tables goes with it.

import { adjustRange, adjustReferenceText } from '../calc/refs';
import { cacheSourceNow, fieldColumnsNow, namedSourceRange, sourceArea, tableState, type PivotCache, type PivotTable, type PivotTableState } from '../model/sheet/pivot';
import { MAX_COLS, MAX_ROWS, foldSheetName, type CellRange, type Sheet, type SheetStructuralEdit, type Workbook } from '../model/sheet/types';

/** What an edit changes of the pivot tables of a workbook (or, applied, what it changed them from): see swapPivotPlaces(). */
export interface PivotPlaces {
  /** By index into Workbook.pivots.caches: the range of the source (PivotCache.sourceRef). */
  caches: { cache: number; sourceRef: string | undefined }[];
  /** By index into Workbook.pivots.tables: where the table stands while not refreshed, and its state after a refresh. */
  tables: { table: number; placed: CellRange | undefined; now: PivotTableState | undefined }[];
}

/**
 * Whether rows and columns may move in a workbook with pivot tables: the model holds every table
 * and every cache of the package, and the source of each is one range or a defined name.
 */
export function pivotsFollow(workbook: Workbook): boolean {
  const pivots = workbook.pivots;
  return !!pivots && pivots.fixed === undefined && pivots.caches.every((cache) => cache.placesFixed === undefined);
}

/** Why the pivot tables keep the rows and columns of a workbook where they are; undefined when they follow. */
export function pivotsFixedReason(workbook: Workbook): string | undefined {
  const pivots = workbook.pivots;
  if (!pivots) return 'its pivot tables cannot be read';
  if (pivots.fixed !== undefined) return pivots.fixed;
  for (const cache of pivots.caches) if (cache.placesFixed !== undefined) return cache.placesFixed;
  return undefined;
}

/** The rows or the columns an edit is about, of a range: [first, last]. */
const span = (range: CellRange, edit: SheetStructuralEdit): [number, number] => (edit.axis === 'row' ? [range.top, range.bottom] : [range.left, range.right]);

/** The cells a table takes on its sheet: its body, and the fields of the page above it with the blank row under them. */
function tableArea(table: PivotTable, body: CellRange): CellRange {
  const pages = table.pages;
  if (!pages) return body;
  // The fields stand in pairs of cells (caption, item) side by side, a blank column between the pairs.
  return { top: Math.max(0, body.top - pages.rows - 1), bottom: body.bottom, left: body.left, right: Math.min(MAX_COLS - 1, Math.max(body.right, body.left + 3 * pages.cols - 2)) };
}

/** The name a refusal gives a cache: the first table made from it. */
function cacheName(workbook: Workbook, cache: PivotCache): string {
  const pivots = workbook.pivots;
  const table = pivots ? pivots.tables.find((one) => one.cache === cache) : undefined;
  return table ? table.name : '';
}

const plural = (edit: SheetStructuralEdit) => (edit.axis === 'row' ? 'rows' : 'cols');

/**
 * Whether a deletion of the cells of a source (`area` on `sheet`, where it stands now) is one the
 * cache can take: throws RangeError when it would take the header row, the column of a field, or
 * all of the source. Which columns are fields is read from the header row as it is now
 * (fieldColumnsNow()): a column whose header names no field (one inserted inside the source) may
 * go, the source gets narrower; when the headers do not match the fields one to one, no column may.
 */
function checkSource(workbook: Workbook, cache: PivotCache, sheet: Sheet, area: CellRange, edit: SheetStructuralEdit): void {
  if (edit.kind !== 'delete') return;
  const end = edit.at + edit.count - 1;
  const [first, last] = span(area, edit);
  if (edit.at > last || end < first) return;
  const refuse = (what: string): never => {
    throw new RangeError(`Cannot delete ${plural(edit)} of the source of the pivot table "${cacheName(workbook, cache)}": ${what} would go`);
  };
  // Rows below the header may go: the source gets shorter. Its header row may not.
  if (edit.axis === 'row') {
    if (edit.at <= first) refuse('its header row');
    return;
  }
  // A deletion of all of it takes the columns of the fields as well (a cache without fields has every column for one).
  const fields = fieldColumnsNow(cache, sheet, area);
  if (fields === 'all' || fields.some((col) => col >= edit.at && col <= end)) refuse('a column of it');
}

/**
 * What an insertion or a deletion of rows or columns in sheet `sheet` makes of the pivot tables of
 * the workbook: the new values, for swapPivotPlaces(); undefined when it changes nothing of them.
 * Changes nothing. Throws RangeError when the edit cannot be made for them: the model does not hold
 * every pivot table (see pivotsFollow()), the edit cuts through a table, or it takes the header row,
 * a column or all of a source.
 */
export function pivotPlacesAfter(workbook: Workbook, sheet: Sheet, edit: SheetStructuralEdit): PivotPlaces | undefined {
  const pivots = workbook.pivots;
  if (!pivots) return undefined;
  const reason = pivotsFixedReason(workbook);
  if (reason !== undefined) throw new RangeError(`Cannot ${edit.kind} ${plural(edit)} in a workbook with pivot tables: ${reason}`);
  const edited = foldSheetName(sheet.name);
  const out: PivotPlaces = { caches: [], tables: [] };
  const limit = edit.axis === 'row' ? MAX_ROWS : MAX_COLS;

  pivots.caches.forEach((cache, index) => {
    const source = cacheSourceNow(cache);
    if (!source) return;
    if (source.kind === 'name') {
      // The name follows the edit by itself (see the defined names in structural()); what it stands for is checked all the same.
      const found = namedSourceRange(workbook, source);
      if (typeof found === 'string' || foldSheetName(found.sheetName) !== edited) return;
      const area = sourceArea(found.ref);
      if (area) checkSource(workbook, cache, sheet, area, edit);
      return;
    }
    if (foldSheetName(source.sheet) !== edited) return;
    const area = sourceArea(source.ref);
    if (!area) throw new RangeError(`Cannot ${edit.kind} ${plural(edit)}: the source of the pivot table "${cacheName(workbook, cache)}" is not one range of cells`);
    checkSource(workbook, cache, sheet, area, edit);
    const ref = adjustReferenceText(source.ref, source.sheet, { ...edit, sheetName: sheet.name });
    const moved = sourceArea(ref);
    const expected = adjustRange(area, edit);
    // What the text says has to be what the range becomes: anything else is not written.
    if (!moved || !expected || moved.top !== expected.top || moved.bottom !== expected.bottom || moved.left !== expected.left || moved.right !== expected.right) {
      throw new RangeError(`Cannot ${edit.kind} ${plural(edit)}: the source of the pivot table "${cacheName(workbook, cache)}" is not one range of cells`);
    }
    if (ref === source.ref) return;
    const file = cache.source && cache.source.kind === 'range' ? cache.source.ref : undefined;
    out.caches.push({ cache: index, sourceRef: ref === file ? undefined : ref });
  });

  pivots.tables.forEach((table, index) => {
    if (table.sheet !== sheet) return;
    const body = tableState(table).location.ref;
    const [first, last] = span(tableArea(table, body), edit);
    const end = edit.at + edit.count - 1;
    if (edit.at > last) return;
    const cuts = edit.kind === 'insert' ? edit.at > first : end >= first;
    if (cuts) throw new RangeError(`Cannot ${edit.kind} ${plural(edit)} through the pivot table "${table.name}"`);
    // All of it moves: down or right by an insertion before it, up or left by a deletion before it.
    const by = edit.kind === 'insert' ? edit.count : -edit.count;
    const [lo, hi] = span(body, edit);
    if (hi + by >= limit) throw new RangeError(`Cannot ${edit.kind} ${plural(edit)}: the pivot table "${table.name}" would be pushed off the sheet`);
    const moved: CellRange = edit.axis === 'row' ? { ...body, top: lo + by, bottom: hi + by } : { ...body, left: lo + by, right: hi + by };
    const file = table.location.ref;
    const same = (a: CellRange, b: CellRange) => a.top === b.top && a.bottom === b.bottom && a.left === b.left && a.right === b.right;
    // The place without the refresh moves as well: undoing the refresh brings the table of the file back where it now stands.
    const base = table.placed || file;
    const placed: CellRange = edit.axis === 'row' ? { ...base, top: base.top + by, bottom: base.bottom + by } : { ...base, left: base.left + by, right: base.right + by };
    const now = table.now ? { ...table.now, location: { ...table.now.location, ref: moved } } : undefined;
    out.tables.push({ table: index, placed: same(placed, file) ? undefined : placed, now });
  });
  return out.caches.length > 0 || out.tables.length > 0 ? out : undefined;
}

/** Puts the places of `places` into the workbook; returns what they were, for undoing. */
export function swapPivotPlaces(workbook: Workbook, places: PivotPlaces): PivotPlaces {
  const pivots = workbook.pivots;
  if (!pivots) throw new RangeError('The workbook has no pivot tables');
  for (const entry of places.caches) if (!pivots.caches[entry.cache]) throw new RangeError(`No pivot cache at index ${entry.cache}`);
  for (const entry of places.tables) if (!pivots.tables[entry.table]) throw new RangeError(`No pivot table at index ${entry.table}`);
  const back: PivotPlaces = { caches: [], tables: [] };
  for (const entry of places.caches) {
    const cache = pivots.caches[entry.cache];
    back.caches.push({ cache: entry.cache, sourceRef: cache.sourceRef });
    if (entry.sourceRef === undefined) delete cache.sourceRef;
    else cache.sourceRef = entry.sourceRef;
  }
  for (const entry of places.tables) {
    const table = pivots.tables[entry.table];
    back.tables.push({ table: entry.table, placed: table.placed, now: table.now });
    if (entry.placed === undefined) delete table.placed;
    else table.placed = entry.placed;
    if (entry.now === undefined) delete table.now;
    else table.now = entry.now;
  }
  workbook.pivotsModified = true;
  return back;
}

/**
 * Whether inserting the rows or columns a deletion takes, at the same place, leaves a source of
 * the sheet other than it was: a deletion of its last rows or columns, which an insertion right
 * after the source does not grow back.
 */
export function pivotSourceLostBy(workbook: Workbook, sheet: Sheet, edit: SheetStructuralEdit): boolean {
  const pivots = workbook.pivots;
  if (!pivots || edit.kind !== 'delete') return false;
  const edited = foldSheetName(sheet.name);
  for (const cache of pivots.caches) {
    const source = cacheSourceNow(cache);
    let ref: string | undefined;
    if (source && source.kind === 'range' && foldSheetName(source.sheet) === edited) ref = source.ref;
    else if (source && source.kind === 'name') {
      const found = namedSourceRange(workbook, source);
      if (typeof found !== 'string' && foldSheetName(found.sheetName) === edited) ref = found.ref;
    }
    const area = ref === undefined ? undefined : sourceArea(ref);
    if (!area) continue;
    const gone = adjustRange(area, edit);
    const back = gone ? adjustRange(gone, { ...edit, kind: 'insert' }) : undefined;
    if (!back || back.top !== area.top || back.bottom !== area.bottom || back.left !== area.left || back.right !== area.right) return true;
  }
  return false;
}
