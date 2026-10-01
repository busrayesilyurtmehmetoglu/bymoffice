// Pixel sizes and positions of rows and columns.
//
// A sheet always has the full 1,048,576 x 16,384 extent, so positions are
// computed from the default size plus the few rows and columns that differ,
// never from a per-index table.

import { MAX_COLS, MAX_ROWS, type Sheet } from '../model/sheet/types';

/** Width in pixels of the widest digit of the default font (Calibri 11 at 96 dpi). */
export const DEFAULT_DIGIT_WIDTH = 7;
/**
 * Width a file stores for a default column. Excel shows this as "8.43": the
 * stored value includes 5 pixels of padding, the displayed one does not.
 */
export const DEFAULT_COLUMN_WIDTH = 9.140625;
/** Points. */
export const DEFAULT_ROW_HEIGHT = 15;

/** Column width as stored in a file (characters, padding included) -> pixels, the way Excel rounds it. */
export function columnWidthToPixels(width: number, digitWidth = DEFAULT_DIGIT_WIDTH): number {
  if (!(width > 0)) return 0;
  return Math.floor(((256 * Math.min(width, 255) + Math.floor(128 / digitWidth)) / 256) * digitWidth);
}

/** Row height in points -> pixels at 96 dpi. */
export function rowHeightToPixels(points: number): number {
  if (!(points > 0)) return 0;
  return Math.round((Math.min(points, 409.5) * 96) / 72);
}

/** Sizes along one axis: `count` items of a default size, with exceptions. */
export class AxisMetrics {
  readonly count: number;
  readonly defaultSize: number;
  readonly total: number;
  /** Indexes with a non-default size, ascending. */
  private readonly indexes: number[];
  private readonly sizes: number[];
  /** deltas[k] = sum over indexes[0..k] of (size - defaultSize). */
  private readonly deltas: number[];
  /** Runs of consecutive hidden (size 0) items, as inclusive [start, end] pairs in ascending order. */
  private readonly hiddenStarts: number[] = [];
  private readonly hiddenEnds: number[] = [];

  constructor(count: number, defaultSize: number, custom: Iterable<[index: number, size: number]> = []) {
    this.count = count;
    this.defaultSize = defaultSize;
    const entries = [...custom].filter(([index, size]) => index >= 0 && index < count && size !== defaultSize);
    entries.sort((a, b) => a[0] - b[0]);
    this.indexes = [];
    this.sizes = [];
    this.deltas = [];
    let delta = 0;
    for (const [index, size] of entries) {
      const last = this.indexes.length - 1;
      if (last >= 0 && this.indexes[last] === index) {
        // The same index twice: the later entry wins.
        delta += size - this.sizes[last];
        this.sizes[last] = size;
        this.deltas[last] = delta;
        continue;
      }
      delta += size - defaultSize;
      this.indexes.push(index);
      this.sizes.push(size);
      this.deltas.push(delta);
    }
    this.total = count * defaultSize + delta;
    if (defaultSize === 0) return;
    for (let k = 0; k < this.indexes.length; k++) {
      if (this.sizes[k] !== 0) continue;
      const last = this.hiddenEnds.length - 1;
      if (last >= 0 && this.hiddenEnds[last] === this.indexes[k] - 1) this.hiddenEnds[last] = this.indexes[k];
      else {
        this.hiddenStarts.push(this.indexes[k]);
        this.hiddenEnds.push(this.indexes[k]);
      }
    }
  }

  /**
   * The first index at `index` or beyond it in the direction of `step` that is
   * not hidden. Returns -1 or `count` when everything that way is hidden. A run
   * of hidden items is jumped over as a whole, however long it is.
   */
  skipHidden(index: number, step: 1 | -1): number {
    if (index < 0) return -1;
    if (index >= this.count) return this.count;
    // Last run that starts at or before `index`.
    let low = 0;
    let high = this.hiddenStarts.length;
    while (low < high) {
      const mid = (low + high) >> 1;
      if (this.hiddenStarts[mid] <= index) low = mid + 1;
      else high = mid;
    }
    const run = low - 1;
    if (run < 0 || this.hiddenEnds[run] < index) return index;
    return step > 0 ? this.hiddenEnds[run] + 1 : this.hiddenStarts[run] - 1;
  }

  /** Number of custom entries with an index below `index`. */
  private customBefore(index: number): number {
    let low = 0;
    let high = this.indexes.length;
    while (low < high) {
      const mid = (low + high) >> 1;
      if (this.indexes[mid] < index) low = mid + 1;
      else high = mid;
    }
    return low;
  }

  size(index: number): number {
    if (index < 0 || index >= this.count) return 0;
    const k = this.customBefore(index);
    return this.indexes[k] === index ? this.sizes[k] : this.defaultSize;
  }

  /** Start position of `index`; offset(count) is the total size. */
  offset(index: number): number {
    if (index <= 0) return 0;
    if (index >= this.count) return this.total;
    const k = this.customBefore(index);
    return index * this.defaultSize + (k > 0 ? this.deltas[k - 1] : 0);
  }

  /**
   * Index of the item that covers position `px`, clamped to the axis. Items of
   * size 0 (hidden) are never returned unless every remaining item is hidden.
   */
  indexAt(px: number): number {
    if (px <= 0) return this.next(0, 1);
    if (px >= this.total) return this.next(this.count - 1, -1);
    // Last index whose offset is <= px; zero-size items share an offset, the last of them wins.
    let low = 0;
    let high = this.count - 1;
    while (low < high) {
      const mid = (low + high + 1) >> 1;
      if (this.offset(mid) <= px) low = mid;
      else high = mid - 1;
    }
    return low;
  }

  /**
   * First index at or after `index` (step 1) or at or before it (step -1) that
   * is not hidden; when everything that way is hidden, the nearest one the other
   * way; when everything is hidden, `index` itself (clamped).
   */
  next(index: number, step: 1 | -1): number {
    const start = Math.min(Math.max(index, 0), this.count - 1);
    let i = this.skipHidden(start, step);
    if (i < 0 || i >= this.count) i = this.skipHidden(start, step > 0 ? -1 : 1);
    return i < 0 || i >= this.count ? start : i;
  }
}

export interface SheetMetrics {
  rows: AxisMetrics;
  cols: AxisMetrics;
  /** Width of the row number column and height of the column letter row, in pixels. */
  headerWidth: number;
  headerHeight: number;
}

/** Wide enough for the largest row number that can come into view. */
export function headerWidthFor(lastRow: number): number {
  const digits = String(Math.max(lastRow, 1) + 1).length;
  return Math.max(40, digits * 8 + 16);
}

export function createSheetMetrics(sheet: Sheet): SheetMetrics {
  const defaultHeight = rowHeightToPixels(sheet.defaultRowHeight ?? DEFAULT_ROW_HEIGHT) || rowHeightToPixels(DEFAULT_ROW_HEIGHT);
  const defaultWidth = columnWidthToPixels(sheet.defaultColWidth ?? DEFAULT_COLUMN_WIDTH) || columnWidthToPixels(DEFAULT_COLUMN_WIDTH);

  const rowSizes: [number, number][] = [];
  for (const [index, row] of sheet.rows) {
    if (row.hidden) rowSizes.push([index, 0]);
    else if (row.height !== undefined) rowSizes.push([index, rowHeightToPixels(row.height)]);
  }

  const colSizes: [number, number][] = [];
  for (const range of sheet.columns) {
    if (!range.hidden && range.width === undefined) continue;
    const size = range.hidden ? 0 : columnWidthToPixels(range.width as number);
    const last = Math.min(range.max, MAX_COLS - 1);
    for (let col = Math.max(range.min, 0); col <= last; col++) colSizes.push([col, size]);
  }

  return {
    rows: new AxisMetrics(MAX_ROWS, defaultHeight, rowSizes),
    cols: new AxisMetrics(MAX_COLS, defaultWidth, colSizes),
    headerWidth: headerWidthFor(MAX_ROWS - 1),
    headerHeight: 22,
  };
}
