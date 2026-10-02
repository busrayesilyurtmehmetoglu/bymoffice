// Table layout: column widths, cell content, row heights, borders, and rows
// as bands that pagination can place, split between pages and repeat.

import type { Block, Border, CellProps, Table, TableProps, TableRow, TableWidth } from '../../model/doc/types';
import { getPrep, layoutBlocks } from './blocks';
import { bandsHeight, pushShifted, type Band, type BorderLine, type CellWhere, type FieldValues, type FloatSpec, type LayoutContext, type RowRef, type StoryRef } from './context';
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
  /** Continues the cell above: no content of its own. */
  continued: boolean;
  /** Height the content is aligned in when the cell spans rows. */
  room?: number;
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
  /** The height is exact: content that starts below it is not shown, as in Word. */
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
  for (const cell of row.cells) {
    if (cell.fill !== undefined && cell.width > 0) items.push({ type: 'rect', x: cell.x + s, y: s, width: Math.max(0, cell.width - 2 * s), height: Math.max(0, height - 2 * s), fill: cell.fill });
  }
  for (const cell of row.cells) {
    if (cell.continued || !cell.bands.length) continue;
    let y = row.topRoom + cell.padTop + s;
    if (cell.align !== 'top') {
      const room = (cell.room === undefined ? height : cell.room) - row.topRoom - cell.padTop - cell.padBottom - 2 * s;
      const spare = room - bandsHeight(cell.bands);
      if (spare > 0) y += cell.align === 'center' ? spare / 2 : spare;
    }
    const x = cell.x + cell.padLeft + s;
    for (const band of cell.bands) {
      y += band.before;
      if (row.exact && y >= height) break;
      pushShifted(items, band.items, x, y);
      if (band.notes) for (const note of band.notes) (notes || (notes = [])).push(note);
      if (band.floats) for (const float of band.floats) (floats || (floats = [])).push({ ...float, dx: float.dx + x, dy: float.dy + y, inCell: true });
      y += band.height + band.after;
    }
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
    if (!cell.continued && bands.length) {
      const room = budget - cell.padTop - cell.padBottom;
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
    first.push({ ...cell, bands: taken, align: 'top', room: undefined, open: false });
    rest.push({ ...cell, bands: kept, align: 'top', room: undefined });
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

/**
 * Lays out a table of the given width budget. Returns one band per row.
 * `path` is the path of the table itself.
 */
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
  const rows = table.rows;
  if (!rows.length || depth >= MAX_LAYOUT_DEPTH) return [];
  const tprops = ctx.tableProps(table);

  // ---- the grid: where every cell is
  const infos: CellInfo[][] = [];
  let columns = table.grid.length;
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    const list: CellInfo[] = [];
    let column = Math.max(0, row.props.gridBefore || 0);
    for (let c = 0; c < row.cells.length; c++) {
      const where = ctx.cellWhere(table, r, c);
      const props = ctx.cellProps(where);
      const span = Math.max(1, props.gridSpan || 1);
      list.push({ index: c, column, span, props, where, merge: props.vMerge === 'restart' ? 1 : props.vMerge === 'continue' ? 2 : 0 });
      column += span;
    }
    column += Math.max(0, row.props.gridAfter || 0);
    if (column > columns) columns = column;
    infos.push(list);
  }
  if (columns === 0) return [];

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

  const widths: number[] = new Array(columns).fill(0);
  /** Columns whose width comes from their content. */
  const natural: boolean[] = new Array(columns).fill(false);
  const scan = Math.min(rows.length, WIDTH_SCAN_ROWS);
  for (let r = 0; r < scan; r++) {
    for (const info of infos[r]) {
      const w = info.props.width;
      if (info.span === 1 && w && w.type === 'dxa' && w.value > 0 && widths[info.column] === 0) widths[info.column] = twips(w.value);
    }
  }
  // Cells that span columns give what is left of their width to the columns nothing else sizes.
  for (let r = 0; r < scan; r++) {
    for (const info of infos[r]) {
      const w = info.props.width;
      if (info.span === 1 || !w || w.type !== 'dxa' || !(w.value > 0)) continue;
      let known = 0;
      let unknown = 0;
      for (let k = info.column; k < info.column + info.span; k++) {
        if (widths[k] > 0) known += widths[k];
        else if (!(twips(table.grid[k]) > 0)) unknown++;
      }
      const spare = twips(w.value) - known;
      if (unknown === 0 || !(spare > 0)) continue;
      for (let k = info.column; k < info.column + info.span; k++) if (widths[k] === 0 && !(twips(table.grid[k]) > 0)) widths[k] = spare / unknown;
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
        for (let c = 0; c < infos[r].length; c++) prepareAll(ctx, rows[r].cells[c].blocks, story, path.concat([r, c]), infos[r][c].where, depth + 1, fields);
      }
      for (let r = 0; r < scan; r++) {
        const list = infos[r];
        for (let c = 0; c < list.length; c++) {
          const info = list[c];
          if (info.span !== 1 || (widths[info.column] > 0 && !natural[info.column])) continue;
          if (info.merge === 2) continue;
          const pct = info.props.width && info.props.width.type === 'pct' && target !== undefined ? (target * info.props.width.value) / 5000 : 0;
          const content = pct > 0 ? pct : cellContentWidth(ctx, table, r, c, story, path, info.where, depth, fields) + defaultLeft + defaultRight;
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
  // Older documents line the text of the first column up with the margin, not the edge of the table.
  if (ctx.compatMode < 15 && depth === 0) x0 -= defaultLeft;
  const align = tprops.align;
  if (align === 'center') x0 = (available - total) / 2;
  else if (align === 'right' || align === 'end') x0 = available - total;

  if (box) {
    box.left = x0;
    box.width = total;
  }

  const spacing = tprops.cellSpacing && tprops.cellSpacing.type === 'dxa' ? Math.min(Math.max(0, twips(tprops.cellSpacing.value)), 48) : 0;
  const tableFill = ctx.shade(tprops.shading);

  /** The cell of a row that covers a grid column. */
  const cellAt = (r: number, column: number): CellInfo | undefined => {
    const list = infos[r];
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
      spacing > 0 ? false : side === 'top' ? r === 0 : side === 'bottom' ? r === rows.length - 1 : side === 'left' ? info.index === 0 : info.index === infos[r].length - 1;
    if (outer) return general && general[side];
    const inside = side === 'top' || side === 'bottom' ? 'insideH' : 'insideV';
    return (borders && borders[inside]) || (general && general[inside]);
  };

  // ---- cells and rows
  const shared: TableShared = { ref: { headers: [] } };
  const boxes: RowBox[] = [];
  /** For each open vertical merge, by first grid column: the row it started in and its cell. */
  const merges = new Map<number, { row: number; cell: CellBox; height: number }>();
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    const list = infos[r];
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
      const box: CellBox = {
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
      if (fill !== undefined) box.fill = fill;
      else if (tableFill !== undefined) box.fill = tableFill;

      // Borders: the top and left edge against the neighbour there, the right edge of the last cell, the bottom edge for the foot of the table.
      const above = r > 0 ? cellAt(r - 1, info.column) : undefined;
      const below = !last ? cellAt(r + 1, info.column) : undefined;
      box.open = below !== undefined && below.merge === 2 && below.column === info.column && info.merge !== 0;
      box.closed = last || !below || spacing > 0;
      if (spacing > 0) {
        box.top = ctx.border(own(r, info, 'top', rprops));
        box.left = ctx.border(own(r, info, 'left', rprops));
        box.right = ctx.border(own(r, info, 'right', rprops));
        box.bottom = ctx.border(own(r, info, 'bottom', rprops));
      } else {
        if (info.merge !== 2) box.top = ctx.border(stronger(own(r, info, 'top', rprops), above ? own(r - 1, above, 'bottom', ctx.tableProps(table, r - 1)) : undefined));
        box.left = ctx.border(stronger(own(r, info, 'left', rprops), c > 0 ? own(r, list[c - 1], 'right', rprops) : undefined));
        if (c === list.length - 1) box.right = ctx.border(own(r, info, 'right', rprops));
        if (!box.open) box.bottom = ctx.border(stronger(own(r, info, 'bottom', rprops), below ? own(r + 1, below, 'top', ctx.tableProps(table, r + 1)) : undefined));
      }
      if (box.top && box.top.width > topRoom) topRoom = box.top.width;
      if (last && box.bottom && box.bottom.width > bottomRoom) bottomRoom = box.bottom.width;

      if (info.merge !== 2) {
        const inner = Math.max(1, box.width - box.padLeft - padRight - 2 * spacing);
        box.bands = layoutBlocks(ctx, row.cells[c].blocks, inner, story, path.concat([r, c]), info.where, depth + 1, fields);
        const height = bandsHeight(box.bands) + box.padTop + box.padBottom;
        if (info.merge === 1) merges.set(info.column, { row: r, cell: box, height });
        else {
          merges.delete(info.column);
          if (height > content) content = height;
        }
      }
      cells.push(box);
    }
    const stated = Math.max(0, twips(row.props.height));
    const exact = row.props.heightRule === 'exact' && stated > 0;
    const min = row.props.heightRule === 'auto' ? 0 : stated;
    const height = exact ? stated : Math.max(min, content + topRoom + bottomRoom + 2 * spacing);
    boxes.push({ cells, height, min: exact ? stated : min, topRoom, bottomRoom, last, spacing, exact });

    // A merged cell that ends with this row: the row grows when the content needs more than the rows it spans.
    merges.forEach((merge, column) => {
      const below = !last ? cellAt(r + 1, column) : undefined;
      if (below && below.merge === 2 && below.column === column) return;
      let spanned = 0;
      for (let k = merge.row; k <= r; k++) spanned += boxes[k].height;
      const need = merge.height + boxes[merge.row].topRoom + 2 * spacing;
      if (need > spanned && !exact) {
        boxes[r].height += need - spanned;
        spanned = need;
      }
      merge.cell.room = spanned;
      merges.delete(column);
    });
  }

  // ---- bands
  const bands: Band[] = [];
  let headers = 0;
  while (headers < rows.length - 1 && rows[headers].props.header) headers++;
  /** Rows a merged cell with more content than its first row runs through stay together. */
  const tied = new Array<boolean>(rows.length).fill(false);
  for (let r = 0; r < rows.length; r++) {
    for (const cell of boxes[r].cells) {
      if (cell.room === undefined || cell.continued) continue;
      if (bandsHeight(cell.bands) + cell.padTop + cell.padBottom + boxes[r].topRoom <= boxes[r].height + EPSILON) continue;
      let covered = boxes[r].height;
      for (let k = r; k < rows.length - 1 && covered < cell.room - EPSILON; k++) {
        tied[k] = true;
        covered += boxes[k + 1].height;
      }
    }
  }
  for (let r = 0; r < rows.length; r++) {
    const row: TableRow = rows[r];
    const exact = row.props.heightRule === 'exact' && boxes[r].min > 0;
    const band = rowBand(shared, boxes[r], r < headers, !row.props.cantSplit && !exact && r >= headers);
    if (tied[r]) band.keep = true;
    // "Keep with next" on the first paragraph of a row keeps the row with the next one.
    if (!band.keep && r < rows.length - 1) {
      const first = row.cells.length ? row.cells[0].blocks[0] : undefined;
      if (first && first.type === 'paragraph' && ctx.paraProps(first.props, infos[r][0].where).keepNext) band.keep = true;
    }
    bands.push(band);
    if (r < headers) shared.ref.headers.push(band);
  }
  return bands;
}
