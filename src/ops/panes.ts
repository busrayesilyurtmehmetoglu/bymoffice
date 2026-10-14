// Frozen panes: rows at the top and columns at the left that stay in place while the rest of the
// sheet scrolls. A change of them is a change of the view of the sheet ("setPane"): no cell
// changes and nothing is calculated again.

import { MAX_COLS, MAX_ROWS, cellName, type PaneState, type Workbook } from '../model/sheet/types';
import { paneOf, paneProblem, samePane, type Op } from './ops';

/**
 * Freezes the first `rows` rows and the first `cols` columns of a sheet; 0 and 0 unfreezes. The
 * frozen rows and columns are counted from the top and the left of the sheet, whatever was
 * scrolled to (the view shows them from there): a first visible cell of the file goes. Returns
 * no operation when the panes are these already.
 */
export function freezePanes(workbook: Workbook, sheetIndex: number, rows: number, cols: number): Op[] {
  const count = (value: unknown, limit: number) => typeof value === 'number' && value >= 0 && value < limit && Math.floor(value) === value;
  if (!count(rows, MAX_ROWS) || !count(cols, MAX_COLS)) throw new RangeError('The number of frozen rows or columns is not valid');
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) throw new RangeError(`No sheet at index ${sheetIndex}`);
  const now = paneOf(sheet);
  let next: PaneState;
  if (rows === 0 && cols === 0) {
    // Unfrozen panes have nothing to say; the cell the view starts at stays what it is.
    if (now.frozenRows === 0 && now.frozenCols === 0) return [];
    next = { frozenRows: 0, frozenCols: 0 };
    if (now.firstVisibleCell !== undefined) next.firstVisibleCell = now.firstVisibleCell;
  } else next = { frozenRows: rows, frozenCols: cols, topLeftCell: cellName(rows, cols) };
  const problem = paneProblem(workbook, sheetIndex, next);
  if (problem) throw new RangeError(problem);
  if (samePane(now, next)) return [];
  return [{ type: 'setPane', sheet: sheetIndex, pane: next }];
}

/** The rows and columns a sheet has frozen. */
export function frozenPanes(workbook: Workbook, sheetIndex: number): { rows: number; cols: number } {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) throw new RangeError(`No sheet at index ${sheetIndex}`);
  const pane = paneOf(sheet);
  return { rows: pane.frozenRows, cols: pane.frozenCols };
}
