// Tables of a Word binary file -> Table blocks of the document model.
//
// The file has no table objects. A table is a run of paragraphs that say
// they are in a table; a cell ends with a cell mark, a row with a row mark,
// and the row mark carries the properties of its row as table sprms: where
// the edges of the cells are (sprmTDefTable), their borders, shading and
// merging, the height of the row. Rows that follow each other are one table.
// A nested table is a run of paragraphs one level deeper inside a cell.
//
// The model wants a grid shared by all rows. It is made from the edges of
// all rows: every distinct position is a grid line, a cell spans the grid
// columns between its edges, and a row that starts further in or ends sooner
// skips columns before or after its cells.

import type { Block, Border, CellMargins, CellProps, RowProps, Shading, Table, TableBorders, TableCell, TablePosition, TableProps, TableRow, TableWidth } from '../../model/doc/types';
import { i16, u16, u8 } from './bytes';
import { countUnknown, type ParaInfo, type PropContext, type TableSprm } from './props';
import { SPRM_T_DEF_TABLE, brc, brc80, colorRef, shd, shd80 } from './sprm';

/** A row may have 63 cells (MS-DOC 2.6.3, sprmTDefTable). */
const MAX_CELLS = 63;

interface CellDef {
  props: CellProps;
  /** Merged into the cell before it. */
  merged: boolean;
}

/** What the row mark of a table row says. */
export interface RowDef {
  /** Positions of the cell edges in twips from the margin: one more than there are cells; empty when the row does not say. */
  edges: number[];
  cells: CellDef[];
  /** Properties of the table as this row states them. */
  table: TableProps;
  row: RowProps;
  /** Half the space between the text of neighbouring cells: the left and right cell margin of Word 97. */
  gapHalf: number | undefined;
}

const SIDES = ['top', 'left', 'bottom', 'right'] as const;
const TABLE_SIDES = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'] as const;
const ALIGNMENTS = ['left', 'center', 'right'];
const VERTICAL: NonNullable<CellProps['vAlign']>[] = ['top', 'center', 'bottom'];
const ANCHORS_V = ['margin', 'page', 'text'];
const ANCHORS_H = ['text', 'margin', 'page'];

/** A width with its unit (Fts, MS-DOC 2.9.106): 0 none, 1 automatic, 2 fiftieths of a percent, 3 twips. */
function tableWidth(b: Uint8Array, at: number): TableWidth | undefined {
  const unit = u8(b, at);
  const value = i16(b, at + 1);
  if (unit === 1) return { type: 'auto', value: 0 };
  if (unit === 2) return { type: 'pct', value };
  if (unit === 3) return { type: 'dxa', value };
  return undefined;
}

function cellOf(cells: CellDef[], index: number): CellDef {
  while (cells.length <= index) cells.push({ props: {}, merged: false });
  return cells[index];
}

const bordersOf = (props: CellProps): TableBorders => props.borders || (props.borders = {});

/** Reads the table sprms of a row mark, in the order the file has them. */
export function readTap(ctx: PropContext, sprms: TableSprm[]): RowDef {
  const def: RowDef = { edges: [], cells: [], table: {}, row: {}, gapHalf: undefined };
  const { cells, table, row } = def;
  let position: TablePosition | undefined;
  const floating = (): TablePosition => position || (position = {});
  /** Runs `fn` for the cells [first, limit) of an operand that starts with these two indexes. */
  const range = (b: Uint8Array, at: number, fn: (cell: CellDef) => void) => {
    const first = u8(b, at);
    const limit = Math.min(u8(b, at + 1), MAX_CELLS + 1);
    for (let i = first; i < limit; i++) fn(cellOf(cells, i));
  };

  for (const { op, b, at, length } of sprms) {
    switch (op) {
      case SPRM_T_DEF_TABLE: {
        const count = Math.min(u8(b, at), MAX_CELLS);
        if (length < 1 + (count + 1) * 2) break;
        def.edges = [];
        cells.length = 0;
        for (let i = 0; i <= count; i++) {
          let edge = i16(b, at + 1 + i * 2);
          // Edges go from left to right; one that does not is pulled up to the edge before it.
          if (i > 0 && edge < def.edges[i - 1]) edge = def.edges[i - 1];
          def.edges.push(edge);
        }
        const first = at + 1 + (count + 1) * 2;
        for (let i = 0; i < count; i++) {
          const cell = cellOf(cells, i);
          const tc = first + i * 20;
          // The cell descriptions may stop short of the last cell; the rest is as Word assumes.
          if (tc + 20 > at + length) continue;
          const flags = u16(b, tc);
          if (flags & 0x0002) cell.merged = true;
          else if (flags & 0x0001) cell.props.hMerge = 'restart';
          if (flags & 0x0004) cell.props.textDirection = flags & 0x0008 ? 'btLr' : flags & 0x0010 ? 'tbRlV' : 'tbRl';
          if (flags & 0x0020) cell.props.vMerge = flags & 0x0040 ? 'restart' : 'continue';
          const vertical = VERTICAL[(flags >> 7) & 3];
          if (vertical !== undefined && vertical !== 'top') cell.props.vAlign = vertical;
          if (flags & 0x1000) cell.props.fitText = true;
          if (flags & 0x2000) cell.props.noWrap = true;
          if (flags & 0x4000) cell.props.hideMark = true;
          for (let k = 0; k < 4; k++) {
            const border = brc80(b, tc + 4 + k * 4, true);
            if (border) bordersOf(cell.props)[SIDES[k]] = border;
          }
        }
        break;
      }
      case 0x5400:
      case 0x548a: {
        const name = ALIGNMENTS[u16(b, at)];
        if (name !== undefined) {
          table.align = name;
          row.align = name;
        }
        break;
      }
      case 0x9601: {
        // sprmTDxaLeft: moves the row so that its first edge is here.
        if (!def.edges.length) def.edges.push(0);
        const shift = i16(b, at) - def.edges[0];
        for (let i = 0; i < def.edges.length; i++) def.edges[i] += shift;
        break;
      }
      case 0x7621: {
        // sprmTInsert: new cells of one width at an index; what follows moves right. Rows of some writers are
        // made of these alone, a cell at a time, instead of a table definition.
        const first = u8(b, at);
        const width = Math.max(0, i16(b, at + 2));
        const edges = def.edges;
        if (first > MAX_CELLS) break;
        if (!edges.length) edges.push(0);
        // An index beyond the last cell: the cells in between come into being too.
        while (edges.length - 1 < first) edges.push(edges[edges.length - 1] + width);
        const count = Math.min(u8(b, at + 1), MAX_CELLS - (edges.length - 1));
        if (count <= 0) break;
        for (let i = first + 1; i < edges.length; i++) edges[i] += count * width;
        const added: number[] = [];
        const fresh: CellDef[] = [];
        for (let i = 1; i <= count; i++) {
          added.push(edges[first] + i * width);
          fresh.push({ props: {}, merged: false });
        }
        edges.splice(first + 1, 0, ...added);
        cellOf(cells, first - 1);
        cells.splice(first, 0, ...fresh);
        break;
      }
      case 0x7623: {
        // sprmTDxaCol: the width of a range of cells; what follows moves with them.
        const edges = def.edges;
        const width = Math.max(0, i16(b, at + 2));
        const limit = Math.min(u8(b, at + 1), edges.length - 1);
        const first = u8(b, at);
        if (first >= limit) break;
        // In one pass: the cells of the range get the width, the edges after them move by what that changed.
        const moved = edges[first] + (limit - first) * width - edges[limit];
        for (let i = first + 1; i <= limit; i++) edges[i] = edges[first] + (i - first) * width;
        for (let k = limit + 1; k < edges.length; k++) edges[k] += moved;
        break;
      }
      case 0x5622: {
        // sprmTDelete: cells are taken out; what follows moves left.
        const edges = def.edges;
        const first = u8(b, at);
        const limit = Math.min(u8(b, at + 1), edges.length - 1);
        if (limit <= first) break;
        const shift = edges[limit] - edges[first];
        for (let k = limit + 1; k < edges.length; k++) edges[k] -= shift;
        edges.splice(first + 1, limit - first);
        cells.splice(first, limit - first);
        break;
      }
      case 0x9602:
        def.gapHalf = Math.max(0, i16(b, at));
        break;
      case 0xd605:
      case 0xd613: {
        const size = op === 0xd605 ? 4 : 8;
        if (length < size * 6) break;
        const borders: TableBorders = {};
        for (let k = 0; k < 6; k++) {
          const border = op === 0xd605 ? brc80(b, at + k * size) : brc(b, at + k * size);
          if (border) borders[TABLE_SIDES[k]] = border;
        }
        table.borders = borders;
        break;
      }
      case 0x9407: {
        // Positive: at least this high; negative: exactly.
        const height = i16(b, at);
        if (height === 0) break;
        row.height = Math.abs(height);
        row.heightRule = height < 0 ? 'exact' : 'atLeast';
        break;
      }
      case 0x3403:
      case 0x3466:
        row.cantSplit = (u8(b, at) & 1) !== 0;
        break;
      case 0x3404:
        row.header = (u8(b, at) & 1) !== 0;
        break;
      case 0xd609: {
        const count = Math.min(Math.floor(length / 2), MAX_CELLS + 1);
        for (let i = 0; i < count; i++) {
          const shading = shd80(b, at + i * 2);
          if (shading) cellOf(cells, i).props.shading = shading;
        }
        break;
      }
      case 0xd612:
      case 0xd616:
      case 0xd60c:
      case 0xd670:
      case 0xd671:
      case 0xd672: {
        // Shading of 22 cells a sprm: the first, second and third sprm continue each other.
        const first = op === 0xd612 || op === 0xd670 ? 0 : op === 0xd616 || op === 0xd671 ? 22 : 44;
        const count = Math.min(Math.floor(length / 10), 22);
        for (let i = 0; i < count && first + i <= MAX_CELLS; i++) {
          const shading = shd(b, at + i * 10);
          if (shading) cellOf(cells, first + i).props.shading = shading;
        }
        break;
      }
      case 0xd62d: {
        const shading: Shading | undefined = length >= 12 ? shd(b, at + 2) : undefined;
        if (shading) range(b, at, (cell) => (cell.props.shading = shading));
        break;
      }
      case 0x7627: {
        const shading = shd80(b, at + 2);
        if (shading) range(b, at, (cell) => (cell.props.shading = shading));
        break;
      }
      case 0xd660: {
        const shading = length >= 10 ? shd(b, at) : undefined;
        if (shading) table.shading = shading;
        break;
      }
      case 0xd620:
      case 0xd62f: {
        // A border for some sides of a range of cells; the third byte says which sides.
        const wide = op === 0xd62f;
        if (length < (wide ? 11 : 7)) break;
        const sides = u8(b, at + 2);
        const border: Border | undefined = wide ? brc(b, at + 3) : brc80(b, at + 3);
        if (!border) break;
        range(b, at, (cell) => {
          const borders = bordersOf(cell.props);
          if (sides & 0x01) borders.top = border;
          if (sides & 0x02) borders.left = border;
          if (sides & 0x04) borders.bottom = border;
          if (sides & 0x08) borders.right = border;
          if (sides & 0x10) borders.tl2br = border;
          if (sides & 0x20) borders.tr2bl = border;
        });
        break;
      }
      case 0xd61a:
      case 0xd61b:
      case 0xd61c:
      case 0xd61d: {
        // Full colours for the borders the cell descriptions give with a colour index.
        const side = SIDES[op - 0xd61a];
        const count = Math.min(Math.floor(length / 4), MAX_CELLS + 1);
        for (let i = 0; i < count && i < cells.length; i++) {
          const borders = cells[i].props.borders;
          const border = borders && borders[side];
          if (border && border.style !== 'nil' && border.style !== 'none') borders![side] = { ...border, color: colorRef(b, at + i * 4) };
        }
        break;
      }
      case 0xd62b: {
        if (length < 2) break;
        const index = u8(b, at);
        if (index > MAX_CELLS) break;
        const flags = u8(b, at + 1);
        const cell = cellOf(cells, index);
        if (flags & 1) cell.props.vMerge = flags & 2 ? 'restart' : 'continue';
        else delete cell.props.vMerge;
        break;
      }
      case 0xd62c: {
        if (length < 3) break;
        const vertical = VERTICAL[u8(b, at + 2)];
        if (vertical !== undefined) range(b, at, (cell) => (cell.props.vAlign = vertical));
        break;
      }
      case 0x5624: {
        // sprmTMerge: the first cell of the range takes the others in.
        const first = u8(b, at);
        range(b, at, (cell) => (cell.merged = true));
        if (first < cells.length) {
          cells[first].merged = false;
          cells[first].props.hMerge = 'restart';
        }
        break;
      }
      case 0x5625:
        range(b, at, (cell) => {
          cell.merged = false;
          delete cell.props.hMerge;
        });
        break;
      case 0xd632:
      case 0xd634:
      case 0xd633: {
        if (length < 6) break;
        const sides = u8(b, at + 2);
        const width = tableWidth(b, at + 3);
        if (!width || width.type !== 'dxa') break;
        if (op === 0xd633) {
          table.cellSpacing = width;
          break;
        }
        const set = (margins: CellMargins) => {
          if (sides & 0x01) margins.top = width;
          if (sides & 0x02) margins.left = width;
          if (sides & 0x04) margins.bottom = width;
          if (sides & 0x08) margins.right = width;
        };
        if (op === 0xd634) set(table.cellMargins || (table.cellMargins = {}));
        else range(b, at, (cell) => set(cell.props.margins || (cell.props.margins = {})));
        break;
      }
      case 0xf614: {
        const width = tableWidth(b, at);
        if (width) table.width = width;
        break;
      }
      case 0xf617: {
        const width = tableWidth(b, at);
        if (width) row.widthBefore = width;
        break;
      }
      case 0xf618: {
        const width = tableWidth(b, at);
        if (width) row.widthAfter = width;
        break;
      }
      case 0x3615:
        table.layout = u8(b, at) & 1 ? 'autofit' : 'fixed';
        break;
      case 0xd635: {
        if (length < 5) break;
        const width = tableWidth(b, at + 2);
        if (width) range(b, at, (cell) => (cell.props.width = width));
        break;
      }
      case 0x563a: {
        const id = ctx.styleId(u16(b, at));
        if (id !== undefined) table.style = id;
        break;
      }
      case 0xd639:
        if (length >= 3) range(b, at, (cell) => (cell.props.noWrap = (u8(b, at + 2) & 1) !== 0));
        break;
      case 0xf636:
        range(b, at, (cell) => (cell.props.fitText = (u8(b, at + 2) & 1) !== 0));
        break;
      case 0xd642:
        if (length >= 3) range(b, at, (cell) => (cell.props.hideMark = (u8(b, at + 2) & 1) !== 0));
        break;
      case 0x7629: {
        const flow = u16(b, at + 2);
        const name = flow === 1 ? 'tbRl' : flow === 3 ? 'btLr' : flow === 4 ? 'lrTbV' : flow === 5 ? 'tbRlV' : undefined;
        range(b, at, (cell) => {
          if (name === undefined) delete cell.props.textDirection;
          else cell.props.textDirection = name;
        });
        break;
      }
      case 0x560b:
      case 0x5664:
        table.bidiVisual = (u16(b, at) & 1) !== 0;
        break;
      case 0x3465:
        if (u8(b, at) & 1) table.overlap = 'never';
        break;
      // Where a table that text flows around is placed.
      case 0x360d: {
        const value = u8(b, at);
        floating().vertAnchor = ANCHORS_V[(value >> 4) & 3] || 'text';
        floating().horzAnchor = ANCHORS_H[(value >> 6) & 3] || 'text';
        break;
      }
      case 0x940e: {
        const value = i16(b, at);
        const target = floating();
        if (value === -4) target.xAlign = 'center';
        else if (value === -8) target.xAlign = 'right';
        else if (value === -12) target.xAlign = 'inside';
        else if (value === -16) target.xAlign = 'outside';
        else target.x = value;
        break;
      }
      case 0x940f: {
        const value = i16(b, at);
        const target = floating();
        if (value === -4) target.yAlign = 'top';
        else if (value === -8) target.yAlign = 'center';
        else if (value === -12) target.yAlign = 'bottom';
        else if (value === -16) target.yAlign = 'inside';
        else if (value === -20) target.yAlign = 'outside';
        else target.y = value;
        break;
      }
      case 0x9410:
        floating().leftFromText = Math.max(0, i16(b, at));
        break;
      case 0x9411:
        floating().topFromText = Math.max(0, i16(b, at));
        break;
      case 0x941e:
        floating().rightFromText = Math.max(0, i16(b, at));
        break;
      case 0x941f:
        floating().bottomFromText = Math.max(0, i16(b, at));
        break;
      case 0x740a: // sprmTTlp: the "auto format" the table was made with
      case 0x7479: // sprmTRsid
      case 0xf661: // sprmTWidthIndent: the edges say the same
        break;
      default:
        countUnknown(ctx, op);
    }
  }
  if (position) table.floating = position;
  return def;
}

const EMPTY_DEF: RowDef = { edges: [], cells: [], table: {}, row: {}, gapHalf: undefined };

interface PendingRow {
  def: RowDef;
  cells: Block[][];
}

const hasText = (blocks: Block[]): boolean => blocks.some((block) => block.type !== 'paragraph' || block.children.length > 0);

/** What rows that state the same may share, compared as text. */
const tableKey = (props: TableProps): string => JSON.stringify([props.borders, props.cellMargins, props.cellSpacing, props.shading]);

/**
 * Left cell margin a table style gives its tables, in twips; undefined when
 * the style (or the default table style, for an undefined id) states none.
 */
export type StyleMargin = (style: string | undefined) => number | undefined;

/** Builds a table from its rows. */
function buildTable(pending: PendingRow[], styleMargin: StyleMargin): Table {
  // The grid: every position an edge of any row is at.
  const positions = new Set<number>();
  /** The edges of each row with every cell at least a twip wide: a cell of no width would have no grid column. */
  const edgesOf = new Map<RowDef, number[]>();
  for (const { def, cells } of pending) {
    let edges = edgesOf.get(def);
    if (!edges) {
      edges = def.edges.slice();
      for (let i = 1; i < edges.length; i++) if (edges[i] <= edges[i - 1]) edges[i] = edges[i - 1] + 1;
      edgesOf.set(def, edges);
    }
    const count = Math.min(edges.length, cells.length + 1);
    for (let i = 0; i < count; i++) positions.add(edges[i]);
  }
  const lines = Array.from(positions).sort((a, b) => a - b);
  const indexOf = new Map<number, number>();
  lines.forEach((position, i) => indexOf.set(position, i));
  const grid: number[] = [];
  for (let i = 1; i < lines.length; i++) grid.push(lines[i] - lines[i - 1]);

  const first = pending[0].def;
  const props: TableProps = { ...first.table };
  const marginsOf = (def: RowDef): CellMargins | undefined => {
    if (def.table.cellMargins && (def.table.cellMargins.left || def.table.cellMargins.right)) return def.table.cellMargins;
    if (def.gapHalf === undefined) return def.table.cellMargins;
    const side: TableWidth = { type: 'dxa', value: def.gapHalf };
    return { ...def.table.cellMargins, left: side, right: side };
  };
  const margins = marginsOf(first);
  if (margins) props.cellMargins = margins;
  if (lines.length) {
    // The model measures the indent to the text of the first cell, as documents of this age do: the edge plus the
    // left cell margin in effect, which is the row's own, that of the table style, or what Word assumes (0.075 inch).
    let left = margins && margins.left && margins.left.type === 'dxa' ? margins.left.value : undefined;
    if (left === undefined) left = styleMargin(props.style);
    props.indent = { type: 'dxa', value: lines[0] + (left === undefined ? 108 : left) };
  }
  const key = tableKey(props);

  const rows: TableRow[] = [];
  for (const { def, cells } of pending) {
    const row: TableRow = { props: { ...def.row }, cells: [] };
    const edges = edgesOf.get(def) || def.edges;
    const known = edges.length >= 2;
    let last: TableCell | undefined;
    for (let i = 0; i < cells.length; i++) {
      const cellDef: CellDef | undefined = def.cells[i];
      const cellProps: CellProps = cellDef ? { ...cellDef.props } : {};
      let span = 1;
      if (known && i + 1 < edges.length) {
        const from = indexOf.get(edges[i])!;
        const to = indexOf.get(edges[i + 1])!;
        span = Math.max(1, to - from);
        cellProps.width = { type: 'dxa', value: edges[i + 1] - edges[i] };
      }
      if (cellDef && cellDef.merged && last) {
        // Merged into the cell before it: that cell grows; text the merged cell still holds is kept with it.
        last.props.gridSpan = (last.props.gridSpan || 1) + span;
        if (last.props.width && cellProps.width && last.props.width.type === 'dxa') {
          last.props.width = { type: 'dxa', value: last.props.width.value + cellProps.width.value };
        }
        if (hasText(cells[i])) for (const block of cells[i]) last.blocks.push(block);
        continue;
      }
      delete cellProps.hMerge;
      if (span > 1) cellProps.gridSpan = span;
      last = { props: cellProps, blocks: cells[i] };
      row.cells.push(last);
    }
    if (known && lines.length) {
      const before = indexOf.get(edges[0])!;
      if (before > 0) {
        row.props.gridBefore = before;
        if (!row.props.widthBefore) row.props.widthBefore = { type: 'dxa', value: edges[0] - lines[0] };
      }
      const end = edges[Math.min(edges.length - 1, cells.length)];
      const after = lines.length - 1 - indexOf.get(end)!;
      if (after > 0) {
        row.props.gridAfter = after;
        if (!row.props.widthAfter) row.props.widthAfter = { type: 'dxa', value: lines[lines.length - 1] - end };
      }
    }
    if (def !== first) {
      // What this row says differently about the table goes with the row.
      const own: TableProps = { ...def.table };
      const ownMargins = marginsOf(def);
      if (ownMargins) own.cellMargins = ownMargins;
      if (tableKey(own) !== key) {
        const exception: TableProps = {};
        if (own.borders) exception.borders = own.borders;
        if (own.cellMargins) exception.cellMargins = own.cellMargins;
        if (own.cellSpacing) exception.cellSpacing = own.cellSpacing;
        if (own.shading) exception.shading = own.shading;
        if (Object.keys(exception).length) row.tableProps = exception;
      }
    }
    rows.push(row);
  }
  return { type: 'table', props, grid, rows };
}

interface Level {
  rows: PendingRow[];
  cells: Block[][];
  current: Block[];
}

/**
 * Takes the paragraphs of a story in order and gives back its blocks:
 * paragraphs outside of tables as they are, the others gathered into tables.
 */
export class BlockAssembler {
  readonly blocks: Block[] = [];
  private readonly levels: Level[] = [];
  private readonly defs = new WeakMap<ParaInfo, RowDef>();

  constructor(
    private readonly ctx: PropContext,
    private readonly styleMargin: StyleMargin = () => undefined,
  ) {}

  /** `endChar` is the character that ends the paragraph: a cell mark (7) ends a cell of an outermost table. */
  add(paragraph: Block, info: ParaInfo, endChar: number): void {
    const depth = info.depth;
    while (this.levels.length > depth) this.close();
    while (this.levels.length < depth) this.levels.push({ rows: [], cells: [], current: [] });
    if (depth === 0) {
      this.blocks.push(paragraph);
      return;
    }
    const level = this.levels[depth - 1];
    const rowEnd = depth === 1 ? info.rowEnd : info.innerRowEnd;
    if (rowEnd) {
      // The row mark itself is not content. A cell left open (no cell mark before the row mark) is closed here.
      if (level.current.length) {
        level.cells.push(level.current);
        level.current = [];
      }
      // Rows with the same properties share what was read from them.
      let def = this.defs.get(info);
      if (!def) this.defs.set(info, (def = readTap(this.ctx, info.table)));
      level.rows.push({ def, cells: level.cells });
      level.cells = [];
      return;
    }
    level.current.push(paragraph);
    if (depth === 1 ? endChar === 7 : info.innerCell) {
      level.cells.push(level.current);
      level.current = [];
    }
  }

  /** Ends the innermost open table and puts it where it belongs. */
  private close(): void {
    const level = this.levels.pop()!;
    // Cells without a row mark after them (a file cut short, or damaged): a row of their own, nothing known about it.
    if (level.current.length) level.cells.push(level.current);
    if (level.cells.length) level.rows.push({ def: EMPTY_DEF, cells: level.cells });
    if (!level.rows.length) return;
    const table = buildTable(level.rows, this.styleMargin);
    const outer = this.levels.length ? this.levels[this.levels.length - 1].current : this.blocks;
    outer.push(table);
  }

  finish(): Block[] {
    while (this.levels.length) this.close();
    return this.blocks;
  }
}
