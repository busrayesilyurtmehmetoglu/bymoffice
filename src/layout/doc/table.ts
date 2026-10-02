// Table layout: column widths, cell content, row heights, borders, and rows
// as bands that pagination can place, split between pages and repeat.
//
// A table is laid out a few rows at a time (tableLayout), so that a table of
// ten thousand rows does not hold up the first page: the column widths are
// found first, from the grid and the first rows, and the rows follow as
// pagination asks for them.

import type { Block, Border, CellProps, Table, TableProps, TableRow, TableWidth } from '../../model/doc/types';
import { getPrep, layoutBlocks } from './blocks';
import { bandsHeight, pushShifted, stackBands, type Band, type BorderLine, type CellWhere, type FieldValues, type FloatSpec, type LayoutContext, type RowRef, type StoryRef } from './context';
import { contentWidths } from './prep';
import type { PageItem } from './types';
import { DEFAULT_CELL_MARGIN, EPSILON, MAX_LAYOUT_DEPTH, twips } from './units';

interface CellBox {
  /** Box of the cell, relative to the left edge of the column. */
  x: number;
  width: number;
  padLeft: number;
  padTop: number;
  padBottom: number;
  /** The content shown in this row: all of it, or for a cell that spans rows the lines that start in this row. */
  bands: Band[];
  fill?: string;
  align: 'top' | 'center' | 'bottom';
  top?: BorderLine;
  left?: BorderLine;
  right?: BorderLine;
  /** Drawn under the last row, and under any row the table is divided after. */
  bottom?: BorderLine;
  /** The cell goes on in the row below: no line under it inside the table. */
  open: boolean;
  /** Nothing below draws the line under the cell: it is drawn with the row. */
  closed: boolean;
  /** Continues the cell above. */
  continued: boolean;
  /**
   * Where the first band starts, from the top of the row, when that is not
   * the top of the cell: content of a cell that spans rows, which runs on
   * from the row above and is aligned in the height of all its rows.
   */
  start?: number;
  /** The content goes on in the row below: its last line here may end there. */
  flows?: boolean;
}

interface RowBox {
  cells: CellBox[];
  height: number;
  /** Height the row must have at least. */
  min: number;
  /** Room the border above the row takes. */
  topRoom: number;
  bottomRoom: number;
  last: boolean;
  /** Space around every cell (tblCellSpacing). */
  spacing: number;
  /** The height is exact: lines that do not fit in it are not shown. */
  exact: boolean;
}

interface TableShared {
  ref: RowRef['table'];
}

const lineItem = (x1: number, y1: number, x2: number, y2: number, b: BorderLine): PageItem => ({ type: 'line', x1, y1, x2, y2, color: b.color, width: b.width, style: b.style });

function rowBand(shared: TableShared, row: RowBox, header: boolean, splittable: boolean): Band {
  const items: PageItem[] = [];
  const height = row.height;
  const s = row.spacing;
  let notes: number[] | undefined;
  let floats: FloatSpec[] | undefined;
  let overhang = 0;
  for (const cell of row.cells) {
    if (cell.fill !== undefined && cell.width > 0) items.push({ type: 'rect', x: cell.x + s, y: s, width: Math.max(0, cell.width - 2 * s), height: Math.max(0, height - 2 * s), fill: cell.fill });
  }
  for (const cell of row.cells) {
    if (!cell.bands.length) continue;
    let y: number;
    if (cell.start !== undefined) y = cell.start;
    else {
      y = row.topRoom + cell.padTop + s;
      if (cell.align !== 'top') {
        const room = height - row.topRoom - cell.padTop - cell.padBottom - 2 * s;
        const spare = room - bandsHeight(cell.bands);
        if (spare > 0) y += cell.align === 'center' ? spare / 2 : spare;
      }
    }
    const x = cell.x + cell.padLeft + s;
    const clipped = row.exact && !cell.flows;
    let shown = 0;
    let end = 0;
    for (const band of cell.bands) {
      y += band.before;
      // An exact height cuts the content off. There is no clipping on the page, so lines that would run into
      // the row below are left out; the first line is shown even when it is taller than the row.
      if (clipped && (y >= height || (shown > 0 && y + band.height > height + EPSILON))) break;
      pushShifted(items, band.items, x, y);
      if (band.notes) for (const note of band.notes) (notes || (notes = [])).push(note);
      if (band.floats) for (const float of band.floats) (floats || (floats = [])).push({ ...float, dx: float.dx + x, dy: float.dy + y, inCell: true });
      end = y + band.height;
      y = end + band.after;
      shown++;
    }
    if (cell.flows && end - height > overhang) overhang = end - height;
  }
  const close: PageItem[] = [];
  for (const cell of row.cells) {
    const x0 = cell.x + s;
    const x1 = cell.x + cell.width - s;
    if (cell.top) items.push(lineItem(x0, s + cell.top.width / 2, x1, s + cell.top.width / 2, cell.top));
    if (cell.left) items.push(lineItem(x0, s, x0, height - s, cell.left));
    if (cell.right) items.push(lineItem(x1, s, x1, height - s, cell.right));
    if (cell.bottom) (cell.closed ? items : close).push(lineItem(x0, height - s, x1, height - s, cell.bottom));
  }
  const band: Band = { height, before: 0, after: 0, items, keep: header, row: { table: shared.ref, header } };
  if (close.length) band.row!.close = close;
  if (overhang > EPSILON) band.overhang = overhang;
  if (notes) band.notes = notes;
  if (floats) band.floats = floats;
  if (splittable) band.split = (available) => splitRow(shared, row, available);
  return band;
}

/** Divides a row so that its first part is `available` high: every cell gives the bands that fit. */
function splitRow(shared: TableShared, row: RowBox, available: number): [Band, Band] | undefined {
  const budget = available - row.topRoom - 2 * row.spacing;
  if (!(budget > 0)) return undefined;
  const first: CellBox[] = [];
  const rest: CellBox[] = [];
  let progress = false;
  let left = false;
  let restHeight = 0;
  for (const cell of row.cells) {
    const taken: Band[] = [];
    const kept: Band[] = [];
    const bands = cell.bands;
    if (bands.length) {
      const top = cell.start !== undefined ? cell.start : row.topRoom + cell.padTop + row.spacing;
      const room = available - top - cell.padBottom - row.spacing;
      let used = 0;
      let k = 0;
      let cut = false;
      while (k < bands.length) {
        const band = bands[k];
        const need = band.before + band.height;
        if (used + need <= room + EPSILON) {
          taken.push(band);
          used += need + band.after;
          k++;
          continue;
        }
        const parts = band.split ? band.split(room - used - band.before) : undefined;
        if (parts) {
          taken.push(parts[0]);
          kept.push(parts[1]);
          k++;
          cut = true;
        }
        break;
      }
      // Lines that must stay together go to the next page together, as long as something stays here.
      if (!cut && k < bands.length) {
        let back = taken.length;
        while (back > 0 && taken[back - 1].keep) back--;
        if (back > 0 && back < taken.length) {
          k -= taken.length - back;
          taken.length = back;
        }
      }
      for (; k < bands.length; k++) kept.push(bands[k]);
      if (taken.length) progress = true;
      if (kept.length) left = true;
      restHeight = Math.max(restHeight, bandsHeight(kept) + cell.padTop + cell.padBottom);
    }
    first.push({ ...cell, bands: taken, align: 'top', open: false, flows: false });
    rest.push({ ...cell, bands: kept, align: 'top', start: undefined });
  }
  if (!progress || !left) return undefined;
  const firstRow: RowBox = { ...row, cells: first, height: available, min: 0, bottomRoom: 0, last: false };
  const restMin = Math.max(0, row.min - available);
  const restRow: RowBox = { ...row, cells: rest, min: restMin, height: Math.max(restMin, restHeight + row.topRoom + row.bottomRoom + 2 * row.spacing) };
  const head = rowBand(shared, firstRow, false, false);
  // The lines under the first part close it at the foot of the page.
  if (head.row!.close) {
    for (const item of head.row!.close) head.items.push(item);
    delete head.row!.close;
  }
  return [head, rowBand(shared, restRow, false, true)];
}

const isNone = (border: Border) => border.style === 'nil' || border.style === 'none';
const weight = (border: Border | undefined) => (border === undefined ? -1 : isNone(border) ? 0 : border.size === undefined ? 4 : border.size);
/** The border shown where two cells meet: the heavier one; an edge nobody draws shows nothing. */
const stronger = (a: Border | undefined, b: Border | undefined) => (weight(b) > weight(a) ? b : a);

interface CellInfo {
  /** Index in the row, first grid column and number of columns. */
  index: number;
  column: number;
  span: number;
  props: CellProps;
  where: CellWhere;
  /** 0 none, 1 starts a vertically merged cell, 2 continues one. */
  merge: 0 | 1 | 2;
}

function measure(width: TableWidth | undefined, whole: number): number | undefined {
  if (!width) return undefined;
  if (width.type === 'dxa') return twips(width.value);
  if (width.type === 'pct') return (whole * width.value) / 5000;
  return undefined;
}

/** Widest the content of a cell can be set, px. */
function cellContentWidth(
  ctx: LayoutContext,
  table: Table,
  r: number,
  c: number,
  story: StoryRef,
  path: number[],
  where: CellWhere,
  depth: number,
  fields: FieldValues | undefined,
): number {
  let max = 0;
  const blocks = table.rows[r].cells[c].blocks;
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    if (block.type === 'paragraph') {
      const widths = contentWidths(getPrep(ctx, block, story, path.concat([r, c, i]), where, depth + 1, fields));
      if (widths.max > max) max = widths.max;
    } else if (block.type === 'table') {
      let sum = 0;
      for (const w of block.grid) sum += twips(w);
      if (sum > max) max = sum;
    }
  }
  return max;
}

/**
 * Prepares every paragraph of the blocks, tables in them included, in
 * document order. Measuring content for column widths looks at cells out of
 * order; numbered paragraphs must have their numbers before that.
 */
function prepareAll(ctx: LayoutContext, blocks: Block[], story: StoryRef, path: number[], where: CellWhere | undefined, depth: number, fields: FieldValues | undefined): void {
  if (depth >= MAX_LAYOUT_DEPTH) return;
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    if (block.type === 'paragraph') getPrep(ctx, block, story, path.concat([i]), where, depth, fields);
    else if (block.type === 'table') {
      for (let r = 0; r < block.rows.length; r++) {
        const cells = block.rows[r].cells;
        for (let c = 0; c < cells.length; c++) prepareAll(ctx, cells[c].blocks, story, path.concat([i, r, c]), ctx.cellWhere(block, r, c), depth + 1, fields);
      }
    }
  }
}

/** Rows whose cells are looked at to find column widths; later rows follow the grid found by then. */
const WIDTH_SCAN_ROWS = 2000;
/** Longest line of text set vertically in a row that states no height, px (six inches). */
const MAX_VERTICAL_RUN = 576;

/** Clockwise turn of the text of a cell with the given text direction; 0 for horizontal text. */
function textTurn(direction: string | undefined): 0 | 90 | 270 {
  // Transitional names and those of the strict schema.
  if (direction === 'btLr' || direction === 'lr') return 270;
  if (direction === 'tbRl' || direction === 'rl' || direction === 'tbRlV' || direction === 'rlV' || direction === 'tbLrV' || direction === 'lrV') return 90;
  return 0;
}

/**
 * An item of content laid out along a line of `length` px, turned into a
 * cell: for 270 degrees the line runs up from the foot of the cell and lines
 * stack from the left, for 90 it runs down from the top and lines stack from
 * the right edge, `across` px away. `offset` moves the lines across the cell.
 */
function turned(item: PageItem, turn: 90 | 270, length: number, across: number, offset: number): PageItem {
  /** Where a point `along` the line and `down` the stack of lines comes to lie. */
  const px = (down: number) => (turn === 270 ? offset + down : across - offset - down);
  const py = (along: number) => (turn === 270 ? length - along : along);
  if (item.type === 'text') return { ...item, x: px(item.y), y: py(item.x), rotation: turn };
  if (item.type === 'line') return { ...item, x1: px(item.y1), y1: py(item.x1), x2: px(item.y2), y2: py(item.x2) };
  if (item.type === 'rect') {
    const x1 = px(item.y);
    const x2 = px(item.y + item.height);
    const y1 = py(item.x);
    const y2 = py(item.x + item.width);
    return { ...item, x: Math.min(x1, x2), y: Math.min(y1, y2), width: Math.abs(x2 - x1), height: Math.abs(y2 - y1) };
  }
  // A picture keeps its size and is turned around its centre.
  const cx = px(item.y + item.height / 2);
  const cy = py(item.x + item.width / 2);
  return { ...item, x: cx - item.width / 2, y: cy - item.height / 2, rotation: (item.rotation || 0) + turn };
}

/** A table being laid out: its rows are handed out a few at a time. */
export interface TableLayout {
  /** The bands of the next rows: at least `count` of them unless the table ends before (more when a cell spans rows beyond them). */
  next(count: number): Band[];
  /** Every row has been handed out. */
  readonly done: boolean;
}

/**
 * Starts the layout of a table of the given width budget: finds the column
 * widths. `path` is the path of the table itself.
 */
export function tableLayout(
  ctx: LayoutContext,
  table: Table,
  available: number,
  story: StoryRef,
  path: number[],
  depth: number,
  fields: FieldValues | undefined,
  /** Filled in with where the table sits in the width it was given. */
  box?: { left: number; width: number },
): TableLayout {
  const rows = table.rows;
  if (!rows.length || depth >= MAX_LAYOUT_DEPTH) return { next: () => [], done: true };
  const tprops = ctx.tableProps(table);

  // ---- the grid: where every cell is. Rows are looked at when they are needed.
  const infos: (CellInfo[] | undefined)[] = [];
  const infoOf = (r: number): CellInfo[] => {
    let list = infos[r];
    if (list) return list;
    list = [];
    const row = rows[r];
    let column = Math.max(0, row.props.gridBefore || 0);
    for (let c = 0; c < row.cells.length; c++) {
      const where = ctx.cellWhere(table, r, c);
      const props = ctx.cellProps(where);
      const span = Math.max(1, props.gridSpan || 1);
      list.push({ index: c, column, span, props, where, merge: props.vMerge === 'restart' ? 1 : props.vMerge === 'continue' ? 2 : 0 });
      column += span;
    }
    infos[r] = list;
    return list;
  };
  const scan = Math.min(rows.length, WIDTH_SCAN_ROWS);
  // The number of columns: that of the grid, or of the widest of the rows looked at for the widths.
  let columns = table.grid.length;
  for (let r = 0; r < scan; r++) {
    const list = infoOf(r);
    const end = (list.length ? list[list.length - 1].column + list[list.length - 1].span : Math.max(0, rows[r].props.gridBefore || 0)) + Math.max(0, rows[r].props.gridAfter || 0);
    if (end > columns) columns = end;
  }
  if (columns === 0) return { next: () => [], done: true };

  // ---- column widths
  const fixed = tprops.layout === 'fixed';
  const cellMargins = tprops.cellMargins || {};
  const marginOf = (own: TableWidth | undefined, general: TableWidth | undefined, fallback: number) => {
    const value = own && own.type === 'dxa' ? own.value : general && general.type === 'dxa' ? general.value : fallback;
    return Math.max(0, twips(value));
  };
  const defaultLeft = marginOf(undefined, cellMargins.left, DEFAULT_CELL_MARGIN);
  const defaultRight = marginOf(undefined, cellMargins.right, DEFAULT_CELL_MARGIN);
  const indent = tprops.indent && tprops.indent.type === 'dxa' ? twips(tprops.indent.value) : 0;
  const room = Math.max(1, available - Math.max(0, indent));
  const target = measure(tprops.width, available);
  const spacing = tprops.cellSpacing && tprops.cellSpacing.type === 'dxa' ? Math.min(Math.max(0, twips(tprops.cellSpacing.value)), 48) : 0;

  /** Length of the lines of a cell whose text is set vertically: the height the row states, else that of the text on one line. */
  const verticalLength = (r: number, info: CellInfo, padTop: number, padBottom: number): number => {
    const row = rows[r];
    const stated = row.props.heightRule === 'auto' ? 0 : Math.max(0, twips(row.props.height));
    // UNSURE: Word lets a row without a height grow with vertical text up to the page; the limit here is a guess.
    const length = stated > 0 ? stated - padTop - padBottom - 2 * spacing : Math.min(cellContentWidth(ctx, table, r, info.index, story, path, info.where, depth, fields), MAX_VERTICAL_RUN);
    return Math.max(1, length);
  };

  const widths: number[] = new Array(columns).fill(0);
  /** Columns whose width comes from their content. */
  const natural: boolean[] = new Array(columns).fill(false);
  for (let r = 0; r < scan; r++) {
    for (const info of infoOf(r)) {
      const w = info.props.width;
      if (info.span === 1 && info.column < columns && w && w.type === 'dxa' && w.value > 0 && widths[info.column] === 0) widths[info.column] = twips(w.value);
    }
  }
  // Cells that span columns give what is left of their width to the columns nothing else sizes.
  for (let r = 0; r < scan; r++) {
    for (const info of infoOf(r)) {
      const w = info.props.width;
      if (info.span === 1 || !w || w.type !== 'dxa' || !(w.value > 0)) continue;
      const end = Math.min(info.column + info.span, columns);
      let known = 0;
      let unknown = 0;
      for (let k = info.column; k < end; k++) {
        if (widths[k] > 0) known += widths[k];
        else if (!(twips(table.grid[k]) > 0)) unknown++;
      }
      const spare = twips(w.value) - known;
      if (unknown === 0 || !(spare > 0)) continue;
      for (let k = info.column; k < end; k++) if (widths[k] === 0 && !(twips(table.grid[k]) > 0)) widths[k] = spare / unknown;
    }
  }
  let unsized = 0;
  for (let k = 0; k < columns; k++) {
    if (widths[k] > 0) continue;
    const grid = k < table.grid.length ? twips(table.grid[k]) : 0;
    if (grid > 0) widths[k] = grid;
    else unsized++;
  }
  if (unsized > 0) {
    if (!fixed) {
      // Automatic layout: a column nothing sizes is as wide as its widest content.
      for (let r = 0; r < scan; r++) {
        const list = infoOf(r);
        for (let c = 0; c < list.length; c++) prepareAll(ctx, rows[r].cells[c].blocks, story, path.concat([r, c]), list[c].where, depth + 1, fields);
      }
      for (let r = 0; r < scan; r++) {
        const list = infoOf(r);
        for (let c = 0; c < list.length; c++) {
          const info = list[c];
          if (info.span !== 1 || info.column >= columns || (widths[info.column] > 0 && !natural[info.column])) continue;
          if (info.merge === 2) continue;
          const pct = info.props.width && info.props.width.type === 'pct' && target !== undefined ? (target * info.props.width.value) / 5000 : 0;
          let content: number;
          if (pct > 0) content = pct;
          else if (textTurn(info.props.textDirection)) {
            // Vertical text: the column is as wide as the lines are thick.
            const margins = info.props.margins || {};
            const rowMargins = ctx.tableProps(table, r).cellMargins || cellMargins;
            const length = verticalLength(r, info, marginOf(margins.top, rowMargins.top, 0), marginOf(margins.bottom, rowMargins.bottom, 0));
            content = bandsHeight(layoutBlocks(ctx, rows[r].cells[c].blocks, length, story, path.concat([r, c]), info.where, depth + 1, fields)) + defaultLeft + defaultRight;
          } else content = cellContentWidth(ctx, table, r, c, story, path, info.where, depth, fields) + defaultLeft + defaultRight;
          if (content > widths[info.column]) widths[info.column] = content;
          natural[info.column] = true;
        }
      }
    }
    let sum = 0;
    let empty = 0;
    for (let k = 0; k < columns; k++) {
      if (widths[k] > 0) sum += widths[k];
      else empty++;
    }
    if (empty > 0) {
      // Still nothing: an equal share of what is left.
      const whole = target !== undefined ? target : room;
      const share = Math.max(whole - sum, 0) / empty || whole / columns;
      for (let k = 0; k < columns; k++) if (!(widths[k] > 0)) widths[k] = share;
    }
  }
  let total = 0;
  for (let k = 0; k < columns; k++) total += widths[k];
  let scale = 1;
  if (target !== undefined && target > 0 && Math.abs(total - target) > 1) {
    // Columns sized by their content give way first; without any, all are scaled.
    let flexible = 0;
    for (let k = 0; k < columns; k++) if (natural[k]) flexible += widths[k];
    if (flexible > 0 && total - flexible < target) {
      const factor = (target - (total - flexible)) / flexible;
      for (let k = 0; k < columns; k++) if (natural[k]) widths[k] *= factor;
    } else scale = target / total;
  } else if (target === undefined && !fixed && total > room) scale = room / total;
  if (scale !== 1 && scale > 0) for (let k = 0; k < columns; k++) widths[k] *= scale;
  const offsets: number[] = new Array(columns + 1);
  offsets[0] = 0;
  for (let k = 0; k < columns; k++) offsets[k + 1] = offsets[k] + widths[k];
  total = offsets[columns];

  // ---- where the table sits
  let x0 = indent;
  if (ctx.compatMode < 15) {
    // Before Word 2013 the indent says where the text of the first cell starts, not the edge of the table: the edge
    // is left of it by the left padding in effect for the first cell of the first row (its own, else that of the
    // table). The same holds in a cell and in a text box; real documents of that kind show it.
    // UNSURE: from Word 2013 on (compatibilityMode 15) the indent is taken to be the edge itself, as before this
    // rule was found; no document of that kind with a telling indent was at hand to check it against.
    const first = infoOf(0)[0];
    const firstMargins = (first && first.props.margins) || {};
    const firstRowMargins = ctx.tableProps(table, 0).cellMargins || cellMargins;
    x0 -= first ? marginOf(firstMargins.left, firstRowMargins.left, DEFAULT_CELL_MARGIN) : defaultLeft;
  }
  const align = tprops.align;
  if (align === 'center') x0 = (available - total) / 2;
  else if (align === 'right' || align === 'end') x0 = available - total;

  if (box) {
    box.left = x0;
    box.width = total;
  }

  const tableFill = ctx.shade(tprops.shading);

  /** The cell of a row that covers a grid column. */
  const cellAt = (r: number, column: number): CellInfo | undefined => {
    const list = infoOf(r);
    let lo = 0;
    let hi = list.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const info = list[mid];
      if (column < info.column) hi = mid - 1;
      else if (column >= info.column + info.span) lo = mid + 1;
      else return info;
    }
    return undefined;
  };
  /** The border a cell asks for on one side, before the neighbour has its say. */
  const own = (r: number, info: CellInfo, side: 'top' | 'left' | 'bottom' | 'right', props: TableProps): Border | undefined => {
    const borders = info.props.borders;
    const stated = borders && borders[side];
    if (stated !== undefined) return stated;
    const general = props.borders;
    // With space between the cells every cell has a box of its own, drawn with the inner borders.
    const outer =
      spacing > 0 ? false : side === 'top' ? r === 0 : side === 'bottom' ? r === rows.length - 1 : side === 'left' ? info.index === 0 : info.index === infoOf(r).length - 1;
    if (outer) return general && general[side];
    const inside = side === 'top' || side === 'bottom' ? 'insideH' : 'insideV';
    return (borders && borders[inside]) || (general && general[inside]);
  };

  // ---- cells and rows
  const shared: TableShared = { ref: { headers: [] } };
  const boxes: RowBox[] = [];
  /** For each open vertical merge, by first grid column: the row it started in, its cell there and its cells in the rows after. */
  const merges = new Map<number, { row: number; cell: CellBox; height: number; parts: CellBox[] }>();
  let headers = 0;
  while (headers < rows.length - 1 && rows[headers].props.header) headers++;

  /** Lays out the cells of row r and settles the height of the rows a merged cell ends in. */
  const layoutRow = (r: number): void => {
    const row = rows[r];
    const list = infoOf(r);
    const rprops = ctx.tableProps(table, r);
    const rowMargins = rprops.cellMargins || cellMargins;
    const cells: CellBox[] = [];
    let content = 0;
    let topRoom = 0;
    let bottomRoom = 0;
    const last = r === rows.length - 1;
    for (let c = 0; c < list.length; c++) {
      const info = list[c];
      // A cell that continues a merge nothing above it started is a cell of its own.
      if (info.merge === 2 && !merges.has(info.column)) info.merge = 0;
      const from = Math.min(info.column, columns);
      const to = Math.min(info.column + info.span, columns);
      const margins = info.props.margins || {};
      const cell: CellBox = {
        x: x0 + offsets[from],
        width: offsets[to] - offsets[from],
        padLeft: marginOf(margins.left, rowMargins.left, DEFAULT_CELL_MARGIN),
        padTop: marginOf(margins.top, rowMargins.top, 0),
        padBottom: marginOf(margins.bottom, rowMargins.bottom, 0),
        bands: [],
        align: info.props.vAlign === 'center' ? 'center' : info.props.vAlign === 'bottom' ? 'bottom' : 'top',
        open: false,
        closed: false,
        continued: info.merge === 2,
      };
      const padRight = marginOf(margins.right, rowMargins.right, DEFAULT_CELL_MARGIN);
      const fill = ctx.shade(info.props.shading);
      if (fill !== undefined) cell.fill = fill;
      else if (tableFill !== undefined) cell.fill = tableFill;

      // Borders: the top and left edge against the neighbour there, the right edge of the last cell, the bottom edge for the foot of the table.
      const above = r > 0 ? cellAt(r - 1, info.column) : undefined;
      const below = !last ? cellAt(r + 1, info.column) : undefined;
      cell.open = below !== undefined && below.merge === 2 && below.column === info.column && info.merge !== 0;
      cell.closed = last || !below || spacing > 0;
      if (spacing > 0) {
        cell.top = ctx.border(own(r, info, 'top', rprops));
        cell.left = ctx.border(own(r, info, 'left', rprops));
        cell.right = ctx.border(own(r, info, 'right', rprops));
        cell.bottom = ctx.border(own(r, info, 'bottom', rprops));
      } else {
        if (info.merge !== 2) cell.top = ctx.border(stronger(own(r, info, 'top', rprops), above ? own(r - 1, above, 'bottom', ctx.tableProps(table, r - 1)) : undefined));
        cell.left = ctx.border(stronger(own(r, info, 'left', rprops), c > 0 ? own(r, list[c - 1], 'right', rprops) : undefined));
        if (c === list.length - 1) cell.right = ctx.border(own(r, info, 'right', rprops));
        if (!cell.open) cell.bottom = ctx.border(stronger(own(r, info, 'bottom', rprops), below ? own(r + 1, below, 'top', ctx.tableProps(table, r + 1)) : undefined));
      }
      if (cell.top && cell.top.width > topRoom) topRoom = cell.top.width;
      if (last && cell.bottom && cell.bottom.width > bottomRoom) bottomRoom = cell.bottom.width;

      if (info.merge !== 2) {
        const inner = Math.max(1, cell.width - cell.padLeft - padRight - 2 * spacing);
        const turn = textTurn(info.props.textDirection);
        if (turn) {
          // The lines run along the height of the row; what they add up to goes across the cell.
          const length = verticalLength(r, info, cell.padTop, cell.padBottom);
          const laid = layoutBlocks(ctx, row.cells[c].blocks, length, story, path.concat([r, c]), info.where, depth + 1, fields);
          const flat: PageItem[] = [];
          const thick = stackBands(laid, flat, 0, 0);
          const spare = inner - thick;
          const offset = cell.align === 'top' || !(spare > 0) ? 0 : cell.align === 'center' ? spare / 2 : spare;
          const items: PageItem[] = [];
          for (let i = 0; i < flat.length; i++) items.push(turned(flat[i], turn, length, inner, offset));
          const band: Band = { height: length, before: 0, after: 0, items, keep: false };
          for (const one of laid) if (one.notes) for (const note of one.notes) (band.notes || (band.notes = [])).push(note);
          cell.bands = laid.length ? [band] : [];
          cell.align = 'top';
        } else cell.bands = layoutBlocks(ctx, row.cells[c].blocks, inner, story, path.concat([r, c]), info.where, depth + 1, fields);
        const height = bandsHeight(cell.bands) + cell.padTop + cell.padBottom;
        if (info.merge === 1) merges.set(info.column, { row: r, cell, height, parts: [] });
        else {
          merges.delete(info.column);
          if (height > content) content = height;
        }
      } else {
        const merge = merges.get(info.column);
        if (merge) merge.parts.push(cell);
      }
      cells.push(cell);
    }
    const stated = Math.max(0, twips(row.props.height));
    const exact = row.props.heightRule === 'exact' && stated > 0;
    const min = row.props.heightRule === 'auto' ? 0 : stated;
    const height = exact ? stated : Math.max(min, content + topRoom + bottomRoom + 2 * spacing);
    boxes.push({ cells, height, min: exact ? stated : min, topRoom, bottomRoom, last, spacing, exact });

    // A merged cell that ends with this row: the row grows when the content needs more than the rows it spans,
    // and the content is dealt out to those rows.
    merges.forEach((merge, column) => {
      const below = !last ? cellAt(r + 1, column) : undefined;
      if (below && below.merge === 2 && below.column === column) return;
      merges.delete(column);
      let spanned = 0;
      for (let k = merge.row; k <= r; k++) spanned += boxes[k].height;
      const need = merge.height + boxes[merge.row].topRoom + 2 * spacing;
      if (need > spanned && !exact) {
        boxes[r].height += need - spanned;
        spanned = need;
      }
      const first = merge.cell;
      const bands = first.bands;
      let y = boxes[merge.row].topRoom + first.padTop + spacing;
      if (first.align !== 'top') {
        const spare = spanned - boxes[merge.row].topRoom - first.padTop - first.padBottom - 2 * spacing - bandsHeight(bands);
        if (spare > 0) y += first.align === 'center' ? spare / 2 : spare;
      }
      // Every line goes to the row it starts in, so that the rows can part at a page end, each with its lines.
      let k = merge.row;
      let part = first;
      let rowTop = 0;
      let rowBottom = boxes[k].height;
      let slice: Band[] = [];
      part.start = y;
      for (let i = 0; i < bands.length; i++) {
        const band = bands[i];
        const top = y + band.before;
        while (k < r && top >= rowBottom - EPSILON) {
          part.bands = slice;
          part.flows = true;
          k++;
          part = merge.parts[k - merge.row - 1] || part;
          rowTop = rowBottom;
          rowBottom += boxes[k].height;
          slice = [];
          part.start = y - rowTop;
        }
        slice.push(band);
        y = top + band.height + band.after;
      }
      part.bands = slice;
      part.flows = k < r;
    });
  };

  const rowKeeps = (r: number): boolean => {
    // "Keep with next" on the first paragraph of a row keeps the row with the next one.
    if (r >= rows.length - 1) return false;
    const row: TableRow = rows[r];
    const first = row.cells.length ? row.cells[0].blocks[0] : undefined;
    return first !== undefined && first.type === 'paragraph' && ctx.paraProps(first.props, infoOf(r)[0].where).keepNext === true;
  };

  /** Rows laid out, and rows handed out as bands. */
  let laid = 0;
  let handed = 0;
  const layout: TableLayout & { done: boolean } = {
    done: false,
    next(count: number): Band[] {
      const out: Band[] = [];
      const target = Math.min(rows.length, laid + Math.max(1, count));
      // A cell that spans rows is settled in the row it ends in: the rows go on until none is open.
      while (laid < rows.length && (laid < target || merges.size > 0)) layoutRow(laid++);
      for (; handed < laid; handed++) {
        const r = handed;
        const row = rows[r];
        const exact = row.props.heightRule === 'exact' && boxes[r].min > 0;
        const band = rowBand(shared, boxes[r], r < headers, !row.props.cantSplit && !exact && r >= headers);
        if (!band.keep && rowKeeps(r)) band.keep = true;
        out.push(band);
        if (r < headers) shared.ref.headers.push(band);
      }
      if (handed >= rows.length) layout.done = true;
      return out;
    },
  };
  return layout;
}

/** Lays out a whole table of the given width budget. Returns one band per row. `path` is the path of the table itself. */
export function tableBands(
  ctx: LayoutContext,
  table: Table,
  available: number,
  story: StoryRef,
  path: number[],
  depth: number,
  fields: FieldValues | undefined,
  /** Filled in with where the table sits in the width it was given. */
  box?: { left: number; width: number },
): Band[] {
  return tableLayout(ctx, table, available, story, path, depth, fields, box).next(Infinity);
}
