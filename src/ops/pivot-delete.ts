// Deleting a pivot table: the table goes, with its cache when no other table is made from it, and
// the cells it takes on its sheet (its body, and the fields of the page above it) are cleared.
//
// The cells are cleared as Excel clears them when a whole table is deleted: nothing is left of
// them, neither a value nor a format (UNSURE: not compared with Excel; what a table shows of its
// formats is the table's, by its style and its own <formats>, and goes with it; a cell format that
// was put on a cell of the table by hand is stored on the cell and goes as well). The widths of
// the columns and the heights of the rows stay.
//
// Like the other commands this only builds operations, one undo step: "deletePivot" (the table and
// its cache leave the lists of the workbook; one of the file is noted for the writer, which takes
// its parts out of the package) and "setCells" (the cells). What else refers to the cache or the
// table (a slicer, a timeline, a pivot chart, a part the model holds no table for) refuses the
// deletion, by the same look through the package as a refresh (lookThroughPivots()).

import { MAX_PIVOT_OUTPUT_CELLS, tableState, type PivotTable } from '../model/sheet/pivot';
import type { Cell, CellRange, Workbook } from '../model/sheet/types';
import type { CellChange, Op } from './ops';
import { lookThroughPivots } from './pivots';
import { tableArea } from './pivot-places';
import { deletePivotFormats } from './rules';

/** A deletion that is not made, with the reason. */
export class PivotDeleteError extends RangeError {
  constructor(message: string) {
    super(message);
    this.name = 'PivotDeleteError';
  }
}

const refuse = (table: PivotTable, reason: string): never => {
  throw new PivotDeleteError(`Cannot delete the pivot table "${table.name}": ${reason}`);
};

const intersects = (a: CellRange, b: CellRange) => a.top <= b.bottom && a.bottom >= b.top && a.left <= b.right && a.right >= b.left;

/** The cells a table takes on its sheet now: its body and the fields of its page. */
export function pivotTableArea(table: PivotTable): CellRange {
  return tableArea(table, tableState(table).location.ref);
}

/**
 * The operations that delete the pivot table at `index` (as listPivots() numbers them), as one
 * undo step. Throws RangeError for an index that is no table, and PivotDeleteError (a RangeError)
 * with the reason when it is not deleted here: something else of the workbook refers to its cache
 * or to it, the model does not hold every pivot table of the package (a table of the file), its
 * cache is not listed by the workbook, it stands over another table. Looks through the workbook
 * first (see lookThroughPivots()) unless `look` is false.
 */
export function deletePivot(workbook: Workbook, index: number, look = true): Op[] {
  const pivots = workbook.pivots;
  const table = pivots && typeof index === 'number' && Math.floor(index) === index && index >= 0 ? pivots.tables[index] : undefined;
  if (!pivots || !table) throw new RangeError(`No pivot table at index ${index}`);
  const sheetIndex = workbook.sheets.indexOf(table.sheet);
  if (sheetIndex < 0) refuse(table, 'its sheet is not in the workbook');
  if (table.sheet.kind !== 'worksheet') refuse(table, 'it does not stand on a worksheet');
  if (table.created === undefined) {
    // The parts of the file are taken out by what the model knows of them: every table and cache of the package.
    if (pivots.fixed !== undefined) refuse(table, pivots.fixed);
    if (table.part === '' || table.cache.part === '') refuse(table, 'its parts are not known');
    if (table.cache.cacheId < 0) refuse(table, 'the workbook does not list its cache');
  }
  if (look) lookThroughPivots(workbook);
  // What names the cache by the places of its items, or the table by its name, would be left naming nothing.
  if (table.cache.referred !== undefined) refuse(table, table.cache.referred);
  // So would what names the relationships the writer drops (other than what it takes out with them, see the reader).
  if (table.created === undefined && pivots.deleteProblem) {
    const cacheGoes = !pivots.tables.some((other) => other !== table && other.cache === table.cache);
    const problem = pivots.deleteProblem(workbook, table, cacheGoes);
    if (problem !== undefined) refuse(table, problem);
  }
  const area = pivotTableArea(table);
  for (const other of pivots.tables) {
    if (other !== table && other.sheet === table.sheet && intersects(pivotTableArea(other), area)) refuse(table, `it stands over the pivot table "${other.name}"`);
  }
  // Its cells: every one of the area that the sheet has, read by the rows the sheet has rather than by the rows of the area.
  const changes: CellChange[] = [];
  const sheet = table.sheet;
  const take = (r: number, cells: (Cell | undefined)[]) => {
    const to = Math.min(area.right, cells.length - 1);
    for (let c = area.left; c <= to; c++) if (cells[c] !== undefined) changes.push({ row: r, col: c, cell: undefined });
    if (changes.length > MAX_PIVOT_OUTPUT_CELLS) refuse(table, `it takes more than ${MAX_PIVOT_OUTPUT_CELLS} cells`);
  };
  if (area.bottom - area.top + 1 <= sheet.rows.size) {
    for (let r = area.top; r <= area.bottom; r++) {
      const row = sheet.rows.get(r);
      if (row) take(r, row.cells);
    }
  } else {
    const rows = [...sheet.rows.keys()].filter((r) => r >= area.top && r <= area.bottom).sort((a, b) => a - b);
    for (const r of rows) {
      const row = sheet.rows.get(r);
      if (row) take(r, row.cells);
    }
  }
  const ops: Op[] = [{ type: 'deletePivot', sheet: sheetIndex, table: index }];
  if (changes.length > 0) ops.push({ type: 'setCells', sheet: sheetIndex, cells: changes });
  // Merged cells the table takes a part of are merged no more: the cells are those of no table (a table stands over none; a damaged file may).
  const merges = sheet.merges.filter((merge) => !intersects(merge, area));
  if (merges.length !== sheet.merges.length) ops.push({ type: 'setMerges', sheet: sheetIndex, merges });
  // The conditional formats of the table go with it.
  if (table.conditionalPriorities) {
    try {
      ops.push(...deletePivotFormats(workbook, sheetIndex, table.conditionalPriorities));
    } catch (error) {
      if (!(error instanceof RangeError)) throw error;
      refuse(table, 'a conditional format of it is one the extension list has more to say about');
    }
  }
  return ops;
}

/** Why the pivot table at `index` would not be deleted; undefined when it would. */
export function deletePivotProblem(workbook: Workbook, index: number, look = true): string | undefined {
  try {
    deletePivot(workbook, index, look);
    return undefined;
  } catch (error) {
    if (error instanceof RangeError) return error.message;
    throw error;
  }
}
