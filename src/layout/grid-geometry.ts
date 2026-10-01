// Where cells are on screen: scrolling, frozen panes and hit testing.
//
// All coordinates are CSS pixels relative to the top-left corner of the grid
// canvas. The headers occupy the top and left edges; frozen rows and columns
// follow them and do not move; the rest scrolls by (scrollX, scrollY).

import type { CellRange, Sheet } from '../model/sheet/types';
import type { SheetMetrics } from './sheet-metrics';

export interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface GridLayout {
  metrics: SheetMetrics;
  /** Canvas size. */
  width: number;
  height: number;
  frozenRows: number;
  frozenCols: number;
  /** Pixel size of the frozen rows and columns (headers not included). */
  frozenWidth: number;
  frozenHeight: number;
  /** How far the scrollable area is scrolled; 0 shows the first row and column after the frozen ones. */
  scrollX: number;
  scrollY: number;
}

/** One of up to four regions that scroll together. Indexes are inclusive; `last < first` means empty. */
export interface GridPane {
  /** Clip rectangle of the pane on the canvas. */
  rect: Rect;
  firstRow: number;
  lastRow: number;
  firstCol: number;
  lastCol: number;
  frozenRows: boolean;
  frozenCols: boolean;
}

export type GridHit =
  | { area: 'cell'; row: number; col: number }
  | { area: 'rowHeader'; row: number }
  | { area: 'colHeader'; col: number }
  | { area: 'corner' };

export function createLayout(
  metrics: SheetMetrics,
  view: Pick<Sheet['view'], 'frozenRows' | 'frozenCols'>,
  viewport: { width: number; height: number; scrollX: number; scrollY: number },
): GridLayout {
  const frozenRows = Math.min(Math.max(view.frozenRows, 0), metrics.rows.count - 1);
  const frozenCols = Math.min(Math.max(view.frozenCols, 0), metrics.cols.count - 1);
  const frozenWidth = metrics.cols.offset(frozenCols);
  const frozenHeight = metrics.rows.offset(frozenRows);
  const max = maxScroll({ metrics, width: viewport.width, height: viewport.height, frozenWidth, frozenHeight });
  return {
    metrics,
    width: viewport.width,
    height: viewport.height,
    frozenRows,
    frozenCols,
    frozenWidth,
    frozenHeight,
    scrollX: Math.min(Math.max(viewport.scrollX, 0), max.x),
    scrollY: Math.min(Math.max(viewport.scrollY, 0), max.y),
  };
}

/** Size of the area in which unfrozen cells are visible. */
export function scrollAreaSize(layout: Pick<GridLayout, 'metrics' | 'width' | 'height' | 'frozenWidth' | 'frozenHeight'>): {
  width: number;
  height: number;
} {
  return {
    width: Math.max(layout.width - layout.metrics.headerWidth - layout.frozenWidth, 0),
    height: Math.max(layout.height - layout.metrics.headerHeight - layout.frozenHeight, 0),
  };
}

/** Largest scroll offsets: the last row and column end at the edge of the canvas. */
export function maxScroll(layout: Pick<GridLayout, 'metrics' | 'width' | 'height' | 'frozenWidth' | 'frozenHeight'>): {
  x: number;
  y: number;
} {
  const area = scrollAreaSize(layout);
  return {
    x: Math.max(layout.metrics.cols.total - layout.frozenWidth - area.width, 0),
    y: Math.max(layout.metrics.rows.total - layout.frozenHeight - area.height, 0),
  };
}

/** Left edge of a column on the canvas (it may be outside the visible area). */
export function colX(layout: GridLayout, col: number): number {
  const offset = layout.metrics.cols.offset(col);
  return layout.metrics.headerWidth + (col < layout.frozenCols ? offset : offset - layout.scrollX);
}

/** Top edge of a row on the canvas. */
export function rowY(layout: GridLayout, row: number): number {
  const offset = layout.metrics.rows.offset(row);
  return layout.metrics.headerHeight + (row < layout.frozenRows ? offset : offset - layout.scrollY);
}

export function cellRect(layout: GridLayout, row: number, col: number): Rect {
  return {
    x: colX(layout, col),
    y: rowY(layout, row),
    width: layout.metrics.cols.size(col),
    height: layout.metrics.rows.size(row),
  };
}

/**
 * Rectangle of a range. A range that spans the frozen boundary is measured
 * from its frozen start to its scrolled end, so it shrinks while its scrolled
 * part moves under the frozen area; the width and height never go below 0.
 */
export function rangeRect(layout: GridLayout, range: CellRange): Rect {
  const x = colX(layout, range.left);
  const y = rowY(layout, range.top);
  const right = colX(layout, range.right) + layout.metrics.cols.size(range.right);
  const bottom = rowY(layout, range.bottom) + layout.metrics.rows.size(range.bottom);
  return { x, y, width: Math.max(right - x, 0), height: Math.max(bottom - y, 0) };
}

/** The panes that have any size, frozen corner first, scrolling pane last. */
export function gridPanes(layout: GridLayout): GridPane[] {
  const { metrics, frozenRows, frozenCols, frozenWidth, frozenHeight } = layout;
  const area = scrollAreaSize(layout);
  const left = metrics.headerWidth;
  const top = metrics.headerHeight;

  // Frozen parts are cut off when the canvas is smaller than they are.
  const frozenVisibleWidth = Math.min(frozenWidth, Math.max(layout.width - left, 0));
  const frozenVisibleHeight = Math.min(frozenHeight, Math.max(layout.height - top, 0));

  const scrollFirstCol = metrics.cols.indexAt(frozenWidth + layout.scrollX);
  const scrollLastCol = area.width > 0 ? metrics.cols.indexAt(frozenWidth + layout.scrollX + area.width - 1) : scrollFirstCol - 1;
  const scrollFirstRow = metrics.rows.indexAt(frozenHeight + layout.scrollY);
  const scrollLastRow = area.height > 0 ? metrics.rows.indexAt(frozenHeight + layout.scrollY + area.height - 1) : scrollFirstRow - 1;

  const frozenLastCol = frozenVisibleWidth > 0 ? Math.min(metrics.cols.indexAt(frozenVisibleWidth - 1), frozenCols - 1) : -1;
  const frozenLastRow = frozenVisibleHeight > 0 ? Math.min(metrics.rows.indexAt(frozenVisibleHeight - 1), frozenRows - 1) : -1;

  const panes: GridPane[] = [];
  const add = (rows: boolean, cols: boolean) => {
    const rect: Rect = {
      x: cols ? left : left + frozenWidth,
      y: rows ? top : top + frozenHeight,
      width: cols ? frozenVisibleWidth : area.width,
      height: rows ? frozenVisibleHeight : area.height,
    };
    if (rect.width <= 0 || rect.height <= 0) return;
    panes.push({
      rect,
      firstRow: rows ? 0 : Math.max(scrollFirstRow, frozenRows),
      lastRow: rows ? frozenLastRow : scrollLastRow,
      firstCol: cols ? 0 : Math.max(scrollFirstCol, frozenCols),
      lastCol: cols ? frozenLastCol : scrollLastCol,
      frozenRows: rows,
      frozenCols: cols,
    });
  };
  if (frozenRows > 0 && frozenCols > 0) add(true, true);
  if (frozenRows > 0) add(true, false);
  if (frozenCols > 0) add(false, true);
  add(false, false);
  return panes;
}

/** Row under a canvas y coordinate, or -1 above the cells. */
function rowAt(layout: GridLayout, y: number): number {
  const local = y - layout.metrics.headerHeight;
  if (local < 0) return -1;
  if (local < layout.frozenHeight) return layout.metrics.rows.indexAt(local);
  return Math.max(layout.metrics.rows.indexAt(local + layout.scrollY), layout.frozenRows);
}

function colAt(layout: GridLayout, x: number): number {
  const local = x - layout.metrics.headerWidth;
  if (local < 0) return -1;
  if (local < layout.frozenWidth) return layout.metrics.cols.indexAt(local);
  return Math.max(layout.metrics.cols.indexAt(local + layout.scrollX), layout.frozenCols);
}

/** What is under a point of the canvas. Points outside the canvas clamp to the nearest cell. */
export function hitTest(layout: GridLayout, x: number, y: number): GridHit {
  const row = rowAt(layout, y);
  const col = colAt(layout, x);
  if (row < 0 && col < 0) return { area: 'corner' };
  if (row < 0) return { area: 'colHeader', col };
  if (col < 0) return { area: 'rowHeader', row };
  return { area: 'cell', row, col };
}

/** Scroll offsets that bring a cell fully into view, moving as little as possible. */
export function scrollToReveal(layout: GridLayout, row: number, col: number): { scrollX: number; scrollY: number } {
  const area = scrollAreaSize(layout);
  const max = maxScroll(layout);
  let { scrollX, scrollY } = layout;
  if (col >= layout.frozenCols) {
    const start = layout.metrics.cols.offset(col) - layout.frozenWidth;
    const end = start + layout.metrics.cols.size(col);
    if (start < scrollX || end - start > area.width) scrollX = start;
    else if (end > scrollX + area.width) scrollX = end - area.width;
  }
  if (row >= layout.frozenRows) {
    const start = layout.metrics.rows.offset(row) - layout.frozenHeight;
    const end = start + layout.metrics.rows.size(row);
    if (start < scrollY || end - start > area.height) scrollY = start;
    else if (end > scrollY + area.height) scrollY = end - area.height;
  }
  return { scrollX: Math.min(Math.max(scrollX, 0), max.x), scrollY: Math.min(Math.max(scrollY, 0), max.y) };
}

/** The merged range that contains a cell, if any. */
export function mergeAt(sheet: Sheet, row: number, col: number): CellRange | undefined {
  for (const range of sheet.merges) {
    if (row >= range.top && row <= range.bottom && col >= range.left && col <= range.right) return range;
  }
  return undefined;
}
