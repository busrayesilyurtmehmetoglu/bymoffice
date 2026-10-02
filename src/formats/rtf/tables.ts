// Tables of an RTF file -> Table blocks of the document model.
//
// As in a Word binary file there are no table objects. "\trowd" starts the
// description of a row: where it begins ("\trleft"), the right edge of each
// cell ("\cellxN", an absolute position, with the properties of the cell in
// front of it), borders, padding, height. Paragraphs marked "\intbl" are in a
// cell; "\cell" ends a cell and "\row" a row. The description that counts
// for a row is the one in effect at its "\row": Word writes it before the
// cells and once more after them, other writers only before.
//
// A table inside a cell is one level deeper ("\itapN"); its cells end with
// "\nestcell" and its rows with "\nestrow", which stands in a group
// ("\nesttableprops") that holds the description of that row.
//
// The model wants a grid shared by all rows; it is made from the edges of all
// rows, as for a binary file (see ../doc/tables.ts).

import type { Block, Border, CellMargins, CellProps, RowProps, Shading, Table, TableBorders, TableCell, TablePosition, TableProps, TableRow, TableWidth } from '../../model/doc/types';
import { AUTO_COLOR, borderWith, lookup, shadingPattern, twipsOf, type FormatContext } from './state';

/** Cells one row may have. Word stops at 63; other writers go further. */
export const MAX_CELLS = 512;
/** Tables in tables that are followed; deeper ones are read at this depth. */
export const MAX_TABLE_DEPTH = 32;

interface CellDef {
  props: CellProps;
  /** Merged into the cell before it ("\clmrg"). */
  merged: boolean;
  /** Shading as the control words state it. */
  share: number;
  patternColor: number;
  background: number;
  /** Padding by side (left, top, right, bottom) and whether its unit is twips. */
  padding: (number | undefined)[];
  paddingUnit: (number | undefined)[];
}

/** What is known about a row at its "\row". */
export interface RowDef {
  /** Positions of the cell edges in twips from the margin: one more than there are cells; one entry when the row states no cell. */
  edges: number[];
  cells: { props: CellProps; merged: boolean }[];
  table: TableProps;
  row: RowProps;
  /** Left and right cell margin: "\trpaddl" and "\trpaddr", else "\trgaph". */
  marginLeft: number;
  marginRight: number;
}

function newCell(): CellDef {
  return { props: {}, merged: false, share: 0, patternColor: 0, background: 0, padding: [], paddingUnit: [] };
}

const TABLE_BORDERS: { [word: string]: keyof TableBorders } = lookup({
  trbrdrt: 'top',
  trbrdrl: 'left',
  trbrdrb: 'bottom',
  trbrdrr: 'right',
  trbrdrh: 'insideH',
  trbrdrv: 'insideV',
});

const CELL_BORDERS: { [word: string]: keyof TableBorders } = lookup({
  clbrdrt: 'top',
  clbrdrl: 'left',
  clbrdrb: 'bottom',
  clbrdrr: 'right',
  cldglu: 'tl2br',
  cldgll: 'tr2bl',
});

const DIRECTIONS: { [word: string]: string | undefined } = lookup({
  cltxlrtb: undefined,
  cltxtbrl: 'tbRl',
  cltxbtlr: 'btLr',
  cltxlrtbv: 'lrTbV',
  cltxtbrlv: 'tbRlV',
});

function widthOf(unit: number | undefined, value: number | undefined): TableWidth | undefined {
  if (unit === 1) return { type: 'auto', value: 0 };
  if (unit === 2 && value !== undefined) return { type: 'pct', value };
  if (unit === 3 && value !== undefined) return { type: 'dxa', value: twipsOf(value) };
  return undefined;
}

/** The description of a row while its control words come in. */
export class RowState {
  private left = 0;
  private gapHalf: number | undefined;
  private rights: number[] = [];
  private cells: CellDef[] = [];
  private cell = newCell();
  private table: TableProps = {};
  private row: RowProps = {};
  private padding: (number | undefined)[] = [];
  private paddingUnit: (number | undefined)[] = [];
  private spacing: number | undefined;
  private spacingUnit: number | undefined;
  private widthUnit: number | undefined;
  private widthValue: number | undefined;
  private share = 0;
  private patternColor = 0;
  private background = 0;
  private position: TablePosition | undefined;
  /** Which border the border words describe: of the table or of the cell being described. */
  private target: { cell: boolean; side: keyof TableBorders } | undefined;
  private made: RowDef | undefined;

  constructor(private readonly ctx: FormatContext) {}

  /** "\trowd": everything back to what a row has when nothing is said. */
  reset(): void {
    this.left = 0;
    this.gapHalf = undefined;
    this.rights = [];
    this.cells = [];
    this.cell = newCell();
    this.table = {};
    this.row = {};
    this.padding = [];
    this.paddingUnit = [];
    this.spacing = undefined;
    this.spacingUnit = undefined;
    this.widthUnit = undefined;
    this.widthValue = undefined;
    this.share = 0;
    this.patternColor = 0;
    this.background = 0;
    this.position = undefined;
    this.target = undefined;
    this.made = undefined;
  }

  /** The state as it is, to put back after a description that must not replace it (that of a nested row). */
  save(): unknown {
    const copy: { [key: string]: unknown } = {};
    const self = this as unknown as { [key: string]: unknown };
    for (const key of Object.keys(self)) if (key !== 'ctx') copy[key] = self[key];
    // What a later control word changes in place is copied.
    copy.rights = this.rights.slice();
    copy.cells = this.cells.slice();
    copy.cell = { ...this.cell, props: { ...this.cell.props }, padding: this.cell.padding.slice(), paddingUnit: this.cell.paddingUnit.slice() };
    copy.table = { ...this.table };
    copy.row = { ...this.row };
    copy.padding = this.padding.slice();
    copy.paddingUnit = this.paddingUnit.slice();
    copy.position = this.position ? { ...this.position } : undefined;
    return copy;
  }

  restore(saved: unknown): void {
    const self = this as unknown as { [key: string]: unknown };
    const from = saved as { [key: string]: unknown };
    for (const key of Object.keys(from)) self[key] = from[key];
  }

  /** A border word; false when no border of a table is being described. */
  border(word: string, param: number): boolean {
    const target = this.target;
    if (!target) return false;
    this.made = undefined;
    if (target.cell) {
      const borders: TableBorders = { ...this.cell.props.borders };
      // "No border stated": the border of the table shows, as for a cell that says nothing.
      if (word === 'brdrtbl') delete borders[target.side];
      else {
        const border = borderWith(this.ctx, borders[target.side], word, param);
        if (border) borders[target.side] = border;
      }
      this.cell.props.borders = borders;
    } else {
      const borders: TableBorders = { ...this.table.borders };
      if (word === 'brdrtbl') delete borders[target.side];
      else {
        const border = borderWith(this.ctx, borders[target.side], word, param);
        if (border) borders[target.side] = border;
      }
      this.table.borders = borders;
    }
    return true;
  }

  /** Border words no longer describe a border of this row. */
  release(): void {
    this.target = undefined;
  }

  private floating(): TablePosition {
    return this.position || (this.position = {});
  }

  /** A control word of a row or cell description; false when it is not one. */
  word(word: string, param: number, hasParam: boolean): boolean {
    const tableSide = TABLE_BORDERS[word];
    if (tableSide !== undefined) {
      this.target = { cell: false, side: tableSide };
      return true;
    }
    const cellSide = CELL_BORDERS[word];
    if (cellSide !== undefined) {
      this.target = { cell: true, side: cellSide };
      return true;
    }
    this.made = undefined;
    const cell = this.cell;
    const on = !hasParam || param !== 0;
    switch (word) {
      case 'cellx': {
        this.target = undefined;
        if (this.rights.length >= MAX_CELLS) return true;
        this.rights.push(twipsOf(param));
        this.cells.push(cell);
        this.cell = newCell();
        return true;
      }
      case 'trleft':
        this.left = twipsOf(param);
        return true;
      case 'trgaph':
        this.gapHalf = Math.max(0, twipsOf(param));
        return true;
      case 'trrh':
        // Positive: at least this high; negative: exactly.
        if (param === 0) {
          delete this.row.height;
          delete this.row.heightRule;
        } else {
          this.row.height = Math.abs(twipsOf(param));
          this.row.heightRule = param < 0 ? 'exact' : 'atLeast';
        }
        return true;
      case 'trhdr':
        this.row.header = on;
        return true;
      case 'trkeep':
        this.row.cantSplit = on;
        return true;
      case 'trql':
        this.table.align = 'left';
        this.row.align = 'left';
        return true;
      case 'trqc':
        this.table.align = 'center';
        this.row.align = 'center';
        return true;
      case 'trqr':
        this.table.align = 'right';
        this.row.align = 'right';
        return true;
      case 'rtlrow':
        this.table.bidiVisual = true;
        return true;
      case 'ltrrow':
        delete this.table.bidiVisual;
        return true;
      case 'trautofit':
        this.table.layout = on ? 'autofit' : 'fixed';
        return true;
      case 'trftsWidth':
        this.widthUnit = param;
        return true;
      case 'trwWidth':
        this.widthValue = param;
        return true;
      case 'trpaddl':
      case 'trpaddt':
      case 'trpaddr':
      case 'trpaddb':
        this.padding['ltrb'.indexOf(word.charAt(6))] = Math.max(0, twipsOf(param));
        return true;
      case 'trpaddfl':
      case 'trpaddft':
      case 'trpaddfr':
      case 'trpaddfb':
        this.paddingUnit['ltrb'.indexOf(word.charAt(7))] = param;
        return true;
      case 'trspdl':
      case 'trspdt':
      case 'trspdr':
      case 'trspdb':
        if (this.spacing === undefined || param > this.spacing) this.spacing = Math.max(0, twipsOf(param));
        return true;
      case 'trspdfl':
      case 'trspdft':
      case 'trspdfr':
      case 'trspdfb':
        this.spacingUnit = param;
        return true;
      case 'trshdng':
        this.share = param;
        return true;
      case 'trcfpat':
        this.patternColor = param;
        return true;
      case 'trcbpat':
        this.background = param;
        return true;

      // Where a table that text flows around is placed.
      case 'tphmrg':
        this.floating().horzAnchor = 'margin';
        return true;
      case 'tphpg':
        this.floating().horzAnchor = 'page';
        return true;
      case 'tphcol':
        this.floating().horzAnchor = 'text';
        return true;
      case 'tpvmrg':
        this.floating().vertAnchor = 'margin';
        return true;
      case 'tpvpg':
        this.floating().vertAnchor = 'page';
        return true;
      case 'tpvpara':
        this.floating().vertAnchor = 'text';
        return true;
      case 'tposx':
        this.floating().x = twipsOf(param);
        return true;
      case 'tposnegx':
        this.floating().x = -Math.abs(twipsOf(param));
        return true;
      case 'tposxc':
        this.floating().xAlign = 'center';
        return true;
      case 'tposxl':
        this.floating().xAlign = 'left';
        return true;
      case 'tposxr':
        this.floating().xAlign = 'right';
        return true;
      case 'tposxi':
        this.floating().xAlign = 'inside';
        return true;
      case 'tposxo':
        this.floating().xAlign = 'outside';
        return true;
      case 'tposy':
        this.floating().y = twipsOf(param);
        return true;
      case 'tposnegy':
        this.floating().y = -Math.abs(twipsOf(param));
        return true;
      case 'tposyt':
        this.floating().yAlign = 'top';
        return true;
      case 'tposyc':
        this.floating().yAlign = 'center';
        return true;
      case 'tposyb':
        this.floating().yAlign = 'bottom';
        return true;
      case 'tposyil':
        this.floating().yAlign = 'inline';
        return true;
      case 'tposyin':
        this.floating().yAlign = 'inside';
        return true;
      case 'tposyout':
        this.floating().yAlign = 'outside';
        return true;
      case 'tdfrmtxtLeft':
        this.floating().leftFromText = Math.max(0, twipsOf(param));
        return true;
      case 'tdfrmtxtRight':
        this.floating().rightFromText = Math.max(0, twipsOf(param));
        return true;
      case 'tdfrmtxtTop':
        this.floating().topFromText = Math.max(0, twipsOf(param));
        return true;
      case 'tdfrmtxtBottom':
        this.floating().bottomFromText = Math.max(0, twipsOf(param));
        return true;
      case 'tabsnoovrlp':
        this.table.overlap = 'never';
        return true;

      case 'clmgf':
        cell.props.hMerge = 'restart';
        return true;
      case 'clmrg':
        cell.merged = true;
        return true;
      case 'clvmgf':
        cell.props.vMerge = 'restart';
        return true;
      case 'clvmrg':
        cell.props.vMerge = 'continue';
        return true;
      case 'clvertalt':
        delete cell.props.vAlign;
        return true;
      case 'clvertalc':
        cell.props.vAlign = 'center';
        return true;
      case 'clvertalb':
        cell.props.vAlign = 'bottom';
        return true;
      case 'clcbpat':
        cell.background = param;
        return true;
      case 'clcfpat':
        cell.patternColor = param;
        return true;
      case 'clshdng':
        cell.share = param;
        return true;
      case 'cltxlrtb':
      case 'cltxtbrl':
      case 'cltxbtlr':
      case 'cltxlrtbv':
      case 'cltxtbrlv': {
        const direction = DIRECTIONS[word];
        if (direction === undefined) delete cell.props.textDirection;
        else cell.props.textDirection = direction;
        return true;
      }
      case 'clFitText':
        cell.props.fitText = true;
        return true;
      case 'clNoWrap':
        cell.props.noWrap = true;
        return true;
      case 'clhidemark':
        cell.props.hideMark = true;
        return true;
      // Word writes the left padding as "\clpadt" and the top padding as "\clpadl", and reads them that way.
      case 'clpadt':
        cell.padding[0] = Math.max(0, twipsOf(param));
        return true;
      case 'clpadl':
        cell.padding[1] = Math.max(0, twipsOf(param));
        return true;
      case 'clpadr':
        cell.padding[2] = Math.max(0, twipsOf(param));
        return true;
      case 'clpadb':
        cell.padding[3] = Math.max(0, twipsOf(param));
        return true;
      case 'clpadft':
        cell.paddingUnit[0] = param;
        return true;
      case 'clpadfl':
        cell.paddingUnit[1] = param;
        return true;
      case 'clpadfr':
        cell.paddingUnit[2] = param;
        return true;
      case 'clpadfb':
        cell.paddingUnit[3] = param;
        return true;
    }
    return false;
  }

  /** The row as described now. Rows described alike share the result until a control word changes the description. */
  def(): RowDef {
    if (this.made) return this.made;
    const ctx = this.ctx;
    const edges = [this.left];
    for (const right of this.rights) {
      // Edges go from left to right; one that does not is pulled up to the edge before it.
      const before = edges[edges.length - 1];
      edges.push(right < before ? before : right);
    }
    const sides = ['left', 'top', 'right', 'bottom'] as const;
    const cells = this.cells.map((cell) => {
      const props: CellProps = { ...cell.props };
      if (cell.share > 0 || cell.background !== 0 || cell.patternColor !== 0) {
        const shading: Shading = { pattern: shadingPattern(cell.share), fill: ctx.color(cell.background) || AUTO_COLOR };
        const color = ctx.color(cell.patternColor);
        if (color) shading.color = color;
        props.shading = shading;
      }
      let margins: CellMargins | undefined;
      for (let i = 0; i < 4; i++) {
        const value = cell.padding[i];
        if (value !== undefined && cell.paddingUnit[i] === 3) (margins || (margins = {}))[sides[i]] = { type: 'dxa', value };
      }
      if (margins) props.margins = margins;
      return { props, merged: cell.merged };
    });

    const table: TableProps = { ...this.table };
    const width = widthOf(this.widthUnit, this.widthValue);
    if (width) table.width = width;
    let margins: CellMargins | undefined;
    for (let i = 0; i < 4; i++) {
      // A unit without a value is a padding of nothing: Word leaves out the zero.
      if (this.paddingUnit[i] === 3) (margins || (margins = {}))[sides[i]] = { type: 'dxa', value: this.padding[i] || 0 };
    }
    if (margins) table.cellMargins = margins;
    if (this.spacing !== undefined && this.spacingUnit === 3 && this.spacing > 0) table.cellSpacing = { type: 'dxa', value: this.spacing };
    if (this.share > 0 || this.background !== 0 || this.patternColor !== 0) {
      const shading: Shading = { pattern: shadingPattern(this.share), fill: ctx.color(this.background) || AUTO_COLOR };
      const color = ctx.color(this.patternColor);
      if (color) shading.color = color;
      table.shading = shading;
    }
    if (this.position) table.floating = { ...this.position };
    // The padding at the sides: what the row states for a side, else half the gap between cells, else none.
    const stated = (side: 'left' | 'right'): number => {
      const margin = margins && margins[side];
      return margin ? margin.value : this.gapHalf || 0;
    };
    this.made = { edges, cells, table, row: { ...this.row }, marginLeft: stated('left'), marginRight: stated('right') };
    return this.made;
  }
}

/** A row nothing is known about: the padding Word gives a table that does not say. */
const EMPTY_DEF: RowDef = { edges: [], cells: [], table: {}, row: {}, marginLeft: 108, marginRight: 108 };

interface PendingRow {
  def: RowDef;
  cells: Block[][];
}

/** Properties of the many cells and rows that state nothing: shared, never changed. */
const EMPTY_CELL: CellProps = Object.freeze({});
const EMPTY_ROW: RowProps = Object.freeze({});
const hasKeys = (object: object): boolean => {
  for (const key in object) if (Object.prototype.hasOwnProperty.call(object, key)) return true;
  return false;
};

/** Tables the reader may still open; one object for all stories of a file. */
export interface TableBudget {
  left: number;
  /** A table was not opened because none were left. */
  exceeded: boolean;
}

const hasText = (blocks: Block[]): boolean => blocks.some((block) => block.type !== 'paragraph' || block.children.length > 0);

/** What rows that state the same may share, compared as text. */
const tableKey = (props: TableProps): string => JSON.stringify([props.borders, props.cellMargins, props.cellSpacing, props.shading]);

const marginsMade = new WeakMap<RowDef, CellMargins>();

/** The cell margins of the table a row belongs to: always stated for the sides, so that layout assumes nothing. */
function marginsOf(def: RowDef): CellMargins {
  // Rows described alike share the result.
  let out = marginsMade.get(def);
  if (!out) {
    out = { ...def.table.cellMargins };
    out.left = { type: 'dxa', value: def.marginLeft };
    out.right = { type: 'dxa', value: def.marginRight };
    marginsMade.set(def, out);
  }
  return out;
}

/** Builds a table from its rows. */
function buildTable(pending: PendingRow[]): Table {
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
    if (edges.length < 2) continue;
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
  props.cellMargins = marginsOf(first);
  if (lines.length) {
    // The model measures the indent to the text of the first cell, as documents of this age do: the edge plus the
    // left cell margin.
    props.indent = { type: 'dxa', value: lines[0] + first.marginLeft };
  }
  const key = tableKey(props);

  const rows: TableRow[] = [];
  for (const { def, cells } of pending) {
    const row: TableRow = { props: hasKeys(def.row) ? { ...def.row } : EMPTY_ROW, cells: [] };
    const edges = edgesOf.get(def) || def.edges;
    const known = edges.length >= 2;
    let last: TableCell | undefined;
    for (let i = 0; i < cells.length; i++) {
      const cellDef = def.cells[i];
      // A cell nothing is known about (the row describes fewer cells, or none) shares its properties with all like it.
      const cellProps: CellProps = cellDef ? { ...cellDef.props } : known && i + 1 < edges.length ? {} : EMPTY_CELL;
      let span = 1;
      if (known && i + 1 < edges.length) {
        const from = indexOf.get(edges[i])!;
        const to = indexOf.get(edges[i + 1])!;
        span = Math.max(1, to - from);
        cellProps.width = { type: 'dxa', value: edges[i + 1] - edges[i] };
      }
      if (cellDef && cellDef.merged && last) {
        // Merged into the cell before it: that cell grows; text the merged cell still holds is kept with it.
        if (last.props === EMPTY_CELL) last.props = {};
        last.props.gridSpan = (last.props.gridSpan || 1) + span;
        if (last.props.width && cellProps.width && last.props.width.type === 'dxa') {
          last.props.width = { type: 'dxa', value: last.props.width.value + cellProps.width.value };
        }
        if (hasText(cells[i])) for (const block of cells[i]) last.blocks.push(block);
        continue;
      }
      if (cellProps.hMerge !== undefined) delete cellProps.hMerge;
      if (span > 1) cellProps.gridSpan = span;
      last = { props: cellProps, blocks: cells[i] };
      row.cells.push(last);
    }
    if (known && lines.length) {
      const before = indexOf.get(edges[0])!;
      const end = edges[Math.min(edges.length - 1, cells.length)];
      const after = lines.length - 1 - indexOf.get(end)!;
      if ((before > 0 || after > 0) && row.props === EMPTY_ROW) row.props = {};
      if (before > 0) {
        row.props.gridBefore = before;
        row.props.widthBefore = { type: 'dxa', value: edges[0] - lines[0] };
      }
      if (after > 0) {
        row.props.gridAfter = after;
        row.props.widthAfter = { type: 'dxa', value: lines[lines.length - 1] - end };
      }
    }
    if (def !== first) {
      // What this row says differently about the table goes with the row.
      const own: TableProps = { ...def.table };
      own.cellMargins = marginsOf(def);
      if (tableKey(own) !== key) {
        const exception: TableProps = {};
        // A row without borders of its own in a table whose first row has some: said, so that it shows none.
        if (own.borders || props.borders) exception.borders = own.borders || noBorders(props.borders!);
        if (own.cellMargins) exception.cellMargins = own.cellMargins;
        if (own.cellSpacing) exception.cellSpacing = own.cellSpacing;
        if (own.shading) exception.shading = own.shading;
        if (Object.keys(exception).length) row.tableProps = exception;
      }
    }
    row.cells = row.cells.slice();
    rows.push(row);
  }
  return { type: 'table', props, grid, rows: rows.slice() };
}

const NO_BORDER: Border = Object.freeze({ style: 'nil' });

/** "No border" for every side the given borders name. */
function noBorders(borders: TableBorders): TableBorders {
  const out: TableBorders = {};
  for (const side of Object.keys(borders) as (keyof TableBorders)[]) out[side] = NO_BORDER;
  return out;
}

interface Level {
  rows: PendingRow[];
  cells: Block[][];
  current: Block[];
  /**
   * The description of the next row, from a row mark that came in front of the cells of its row (see endRow);
   * undefined in a table whose row marks follow their cells.
   */
  early: RowDef | undefined;
}

/** A budget that is never used up, for an assembler that is given none. */
const NO_BUDGET: TableBudget = { left: Infinity, exceeded: false };

/**
 * Takes the paragraphs of a story in order and gives back its blocks:
 * paragraphs outside of tables as they are, the others gathered into tables.
 */
export class TableAssembler {
  readonly blocks: Block[] = [];
  private readonly levels: Level[] = [];

  constructor(private readonly budget: TableBudget = NO_BUDGET) {}

  /** Depth of the innermost open table. */
  get depth(): number {
    return this.levels.length;
  }

  /**
   * The table open at a depth, opened now when it is not yet, with the tables around it. When the reader has no
   * tables left to open (see TableBudget) this is the innermost table that is open, and undefined when none is.
   */
  private open(depth: number): Level | undefined {
    while (this.levels.length > depth) this.close();
    const budget = this.budget;
    while (this.levels.length < depth) {
      if (budget.left <= 0) {
        budget.exceeded = true;
        break;
      }
      budget.left--;
      this.levels.push({ rows: [], cells: [], current: [], early: undefined });
    }
    return this.levels.length ? this.levels[this.levels.length - 1] : undefined;
  }

  /** A paragraph (or other block) at a table depth; 0 is outside of every table. */
  add(block: Block, depth: number): void {
    if (depth <= 0) {
      while (this.levels.length) this.close();
      this.blocks.push(block);
      return;
    }
    const level = this.open(Math.min(depth, MAX_TABLE_DEPTH));
    if (level) level.current.push(block);
    else this.blocks.push(block);
  }

  /** Ends the cell that is open at a depth; a cell nothing was put into is a cell all the same. */
  endCell(depth: number): void {
    const level = this.open(Math.max(1, Math.min(depth, MAX_TABLE_DEPTH)));
    if (!level) return;
    level.cells.push(level.current.slice());
    level.current.length = 0;
  }

  /**
   * Ends the row at a depth with the description in effect. `ahead` says that the row mark may stand in front of
   * the cells of its row (it is the one of a "\nesttableprops" group): when a table starts with such a mark, each
   * of its marks describes the cells after it.
   */
  endRow(depth: number, def: RowDef, ahead = false): void {
    const level = this.open(Math.max(1, Math.min(depth, MAX_TABLE_DEPTH)));
    if (!level) return;
    // A cell left open (no cell mark before the row mark) is closed here.
    if (level.current.length) {
      level.cells.push(level.current.slice());
      level.current.length = 0;
    }
    if (!level.cells.length) {
      // A row mark without cells before it is not a row; at the start of a table it may describe the row to come.
      if (ahead && !level.rows.length) level.early = def;
      return;
    }
    if (level.early) {
      level.rows.push({ def: level.early, cells: level.cells });
      level.early = def;
    } else level.rows.push({ def, cells: level.cells });
    level.cells = [];
  }

  /** Ends the innermost open table and puts it where it belongs. */
  private close(): void {
    const level = this.levels.pop()!;
    // Cells without a row mark after them (a file cut short, or damaged): a row of their own, with the description
    // that came ahead of them when there is one, else nothing known about it.
    if (level.current.length) level.cells.push(level.current.slice());
    if (level.cells.length) level.rows.push({ def: level.early || EMPTY_DEF, cells: level.cells });
    if (!level.rows.length) return;
    const table = buildTable(level.rows);
    const outer = this.levels.length ? this.levels[this.levels.length - 1].current : this.blocks;
    outer.push(table);
  }

  finish(): Block[] {
    while (this.levels.length) this.close();
    return this.blocks.slice();
  }
}
