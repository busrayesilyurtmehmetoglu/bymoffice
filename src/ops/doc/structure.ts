// Structural commands: tables and their rows and columns, hyperlinks,
// pictures.

import {
  EMPTY_PARA_PROPS,
  safeHyperlinkTarget,
  type Block,
  type Border,
  type CellProps,
  type DocMedia,
  type DocModel,
  type DrawingItem,
  type Inline,
  type Paragraph,
  type RowProps,
  type RunItem,
  type RunProps,
  type Table,
  type TableCell,
  type TableProps,
  type TableRow,
  type Wrapper,
} from '../../model/doc/types';
import type { XmlElement, XmlNode } from '../../xml';
import { buildEdit, type DocEdit, type DocEditOptions, type Draft } from './draft';
import { patchRunProps } from './format';
import {
  caretIndex,
  cleanText,
  fieldPairs,
  flatten,
  groupFor,
  harvestBlocks,
  harvestRow,
  harvestTable,
  inheritedProps,
  isRemovable,
  itemPiece,
  leftInlines,
  offsetOf,
  paragraphWith,
  rebuild,
  startIndex,
  textPieces,
  wrapBetween,
  type Harvest,
  type Piece,
} from './pieces';
import { asRange, containerBlocks, inTextBox, orderRange, paragraphIn, resolvePosition, storyOf, type DocPosition, type DocRange } from './position';
import { newMarkProps, placeBlock } from './text';
import { TRACK_UNSUPPORTED, insertionWrap, plainProps, revisionLike, strikePieces, type Tracking } from './track';
import { insertedCell, insertedRow, insertedTable, isOwnCell, isOwnRow, struckCell, struckRow } from './track-blocks';

const bad = (message: string): never => {
  throw new RangeError(message);
};

const whole = (value: unknown, min: number, max: number): value is number => typeof value === 'number' && value >= min && value <= max && Math.floor(value) === value;

/** Most columns a table of Word has. */
const MAX_COLUMNS = 63;
/** Most cells a table made in one step has. */
const MAX_NEW_CELLS = 100000;

const spanOf = (props: CellProps): number => (props.gridSpan !== undefined && props.gridSpan > 1 ? props.gridSpan : 1);

/** Grid columns of the cells of a row, as [start, end) each. */
export function rowSpans(row: TableRow): [number, number][] {
  const out: [number, number][] = [];
  let at = row.props.gridBefore || 0;
  for (const cell of row.cells) {
    const span = spanOf(cell.props);
    out.push([at, at + span]);
    at += span;
  }
  return out;
}

/**
 * An empty paragraph formatted like the first paragraph of a cell: what a new cell beside it
 * holds. It is new: of the tracked changes of that paragraph (of its mark, of its formatting) it
 * has none.
 */
function emptyLike(cell: TableCell | undefined): Paragraph {
  if (cell) {
    for (const block of cell.blocks) {
      if (block.type !== 'paragraph') continue;
      return { type: 'paragraph', props: newMarkProps(block.props), children: [] };
    }
  }
  return { type: 'paragraph', props: EMPTY_PARA_PROPS, children: [] };
}

interface CellSite {
  story: string;
  /** Path of the list of blocks the table is in. */
  container: number[];
  index: number;
  table: Table;
  row: number;
  cell: number;
}

/** The innermost table cell a position is in. Throws RangeError when it is in none. */
function cellSite(doc: DocModel, position: DocPosition): CellSite {
  resolvePosition(doc, position);
  const path = position.path;
  if (inTextBox(path)) bad('The content of a text box cannot be edited');
  if (path.length < 4) bad('The position is not in a table');
  const container = path.slice(0, -4);
  const index = path[path.length - 4];
  const table = containerBlocks(doc, position.story, container)[index];
  if (!table || table.type !== 'table') return bad('The position is not in a table');
  return { story: position.story, container, index, table, row: path[path.length - 3], cell: path[path.length - 2] };
}

function tableWith(table: Table, rows: TableRow[], grid = table.grid, props = table.props): Table {
  const out: Table = { ...table, rows, grid, props };
  delete out.srcStart;
  delete out.srcEnd;
  return out;
}

/** Start of the first paragraph of a cell of the table a site names, after the table was replaced. */
function cellStart(doc: DocModel, site: CellSite, row: number, cell: number): DocPosition {
  const base = site.container.concat([site.index, row, cell]);
  const path = paragraphIn(containerBlocks(doc, site.story, base), base, 0, 1);
  if (!path) return bad('The cell holds no paragraph');
  return { story: site.story, path, offset: 0 };
}

/**
 * Inserts a row above or below the row a position is in. The new row has the
 * cells of that row, empty, with their formatting; where the row is crossed
 * by vertically merged cells, the new cells continue the merge. Where changes
 * are tracked the row is a tracked insertion (see track-blocks.ts).
 *
 * Of the tracked changes of the row beside it the new row takes none, but
 * one: a cell in a column that is a pending change (put in or taken out,
 * <w:cellIns> or <w:cellDel>) is a cell of that column, and says so under an
 * id of its own. A column is resolved whole, in every row; a row without the
 * mark would keep a cell the others lose.
 */
export function insertRow(doc: DocModel, position: DocPosition, where: 'above' | 'below', _options?: DocEditOptions): DocEdit {
  if (where !== 'above' && where !== 'below') bad('A row goes above or below');
  return buildEdit(doc, (draft) => {
    const site = cellSite(doc, position);
    const table = site.table;
    const ref = table.rows[site.row];
    const spans = rowSpans(ref);
    const below = table.rows[site.row + 1];
    const belowSpans = below ? rowSpans(below) : [];
    const cells: TableCell[] = ref.cells.map((cell, c) => {
      // A new cell: formatted like the one beside it, without the tracked changes that one records.
      const props: CellProps = { ...plainProps(cell.props) };
      let continues = false;
      if (where === 'above') continues = cell.props.vMerge === 'continue';
      else if (cell.props.vMerge !== undefined && below) {
        for (let k = 0; k < below.cells.length; k++) if (belowSpans[k][0] === spans[c][0] && below.cells[k].props.vMerge === 'continue') continues = true;
      }
      if (continues) props.vMerge = 'continue';
      else delete props.vMerge;
      if (cell.props.inserted) props.inserted = revisionLike(doc, cell.props.inserted);
      if (cell.props.deleted) props.deleted = revisionLike(doc, cell.props.deleted);
      return { props, blocks: [emptyLike(cell)] };
    });
    const props: RowProps = { ...plainProps(ref.props) };
    // Only the rows at the top of a table repeat on every page.
    if (where === 'below') delete props.header;
    let row: TableRow = { props, cells };
    if (ref.tableProps) row.tableProps = plainProps(ref.tableProps);
    const at = where === 'above' ? site.row : site.row + 1;
    const wrap = wrapBetween(at > 0 ? table.rows[at - 1].wrap : undefined, at < table.rows.length ? table.rows[at].wrap : undefined);
    if (wrap) row.wrap = wrap;
    if (draft.track) row = insertedRow(draft.track, row);
    const rows = table.rows.slice();
    rows.splice(at, 0, row);
    draft.replace(site.story, site.container, site.index, 1, [tableWith(table, rows)]);
    return cellStart(doc, site, at, site.cell);
  }, _options, true);
}

/** What removed rows or cells leave behind (see harvestBlocks): inlines for a paragraph that stays, and blocks without content. */
export interface Leftover {
  inlines: Inline[];
  blocks: Block[];
}

const leftoverOf = (harvest: Harvest): Leftover => ({ inlines: leftInlines(harvest.pieces), blocks: harvest.blocks });

/**
 * A cell that takes over what removed content left behind: the inlines go to
 * the start of its first paragraph (`atEnd`: to the end of its last one), the
 * blocks before that paragraph. The cell itself when there is nothing.
 */
export function cellWith(cell: TableCell, left: Leftover, atEnd: boolean): TableCell {
  if (left.inlines.length === 0 && left.blocks.length === 0) return cell;
  const blocks = cell.blocks.slice();
  let index = -1;
  for (let i = atEnd ? blocks.length - 1 : 0; i >= 0 && i < blocks.length && index < 0; i += atEnd ? -1 : 1) if (blocks[i].type === 'paragraph') index = i;
  if (index < 0) {
    index = blocks.length;
    blocks.push({ type: 'paragraph', props: EMPTY_PARA_PROPS, children: left.inlines });
  } else if (left.inlines.length) {
    const paragraph = blocks[index] as Paragraph;
    blocks[index] = paragraphWith(paragraph, atEnd ? paragraph.children.concat(left.inlines) : left.inlines.concat(paragraph.children));
  }
  if (left.blocks.length) blocks.splice(index, 0, ...left.blocks);
  return { ...cell, blocks };
}

/** A row whose first cell (`atEnd`: last cell) takes over what removed content left behind. */
export function rowWith(row: TableRow, left: Leftover, atEnd: boolean): TableRow {
  if (row.cells.length === 0 || (left.inlines.length === 0 && left.blocks.length === 0)) return row;
  const cells = row.cells.slice();
  const c = atEnd ? cells.length - 1 : 0;
  cells[c] = cellWith(cells[c], left, atEnd);
  return { ...row, cells };
}

/**
 * Removes a table; an empty paragraph takes its place where the list of
 * blocks would otherwise end with a table, be empty, or have two tables
 * touch. What the table held that is not content (marks of bookmarks and
 * comments, kept XML) goes into the paragraph after it, else the one before
 * it, else into such a paragraph.
 */
function removeTable(draft: Draft, site: CellSite): DocPosition {
  const blocks = containerBlocks(draft.doc, site.story, site.container);
  const near = (step: 1 | -1): number => {
    for (let i = site.index + step, tries = 0; i >= 0 && i < blocks.length && tries < 64; i += step, tries++) {
      if (blocks[i].type === 'paragraph' || blocks[i].type === 'table') return i;
    }
    return -1;
  };
  const harvest: Harvest = { pieces: [], blocks: [] };
  harvestTable(site.table, harvest);
  const left = leftoverOf(harvest);
  const before = near(-1);
  const after = near(1);
  const tableBefore = before >= 0 && blocks[before].type === 'table';
  const tableAfter = after >= 0 && blocks[after].type === 'table';
  let needed = after >= 0 ? tableAfter && tableBefore : before < 0 || tableBefore;
  if (!needed && left.inlines.length) {
    const target = after >= 0 && !tableAfter ? after : before >= 0 && !tableBefore ? before : -1;
    if (target < 0) needed = true;
    else {
      const paragraph = blocks[target] as Paragraph;
      draft.setParagraph(site.story, site.container.concat([target]), paragraphWith(paragraph, target === after ? left.inlines.concat(paragraph.children) : paragraph.children.concat(left.inlines)));
    }
  }
  const out: Block[] = left.blocks.slice();
  if (needed) {
    // (Where changes are tracked only a table the author put in since the document was opened comes here, and that
    // one stands between the paragraphs it was put between: the list needs no other.)
    const filler: Paragraph = { type: 'paragraph', props: EMPTY_PARA_PROPS, children: left.inlines };
    if (site.table.wrap) filler.wrap = site.table.wrap;
    out.push(filler);
  }
  draft.replace(site.story, site.container, site.index, 1, out);
  const now = containerBlocks(draft.doc, site.story, site.container);
  const path = paragraphIn(now, site.container, site.index, 1) || paragraphIn(now, site.container, site.index - 1, -1);
  if (!path) return bad('The story holds no paragraph');
  const end = path[site.container.length] < site.index;
  return { story: site.story, path, offset: end ? resolvePosition(draft.doc, { story: site.story, path, offset: 0 }).length : 0 };
}

/**
 * Deletes the row a position is in; the last row takes the table with it. A
 * vertically merged cell that began in the row begins in the next row
 * afterwards. What the row held that is not content (marks of bookmarks and
 * comments, kept XML) moves to the start of the row after it, or to the end
 * of the row before it.
 *
 * Where changes are tracked the row stays, as a tracked deletion with
 * everything in it deleted, and the table with it; only a row the same
 * author inserted since the document was opened, with nothing of anybody
 * else in it, goes for good.
 */
export function deleteRow(doc: DocModel, position: DocPosition, _options?: DocEditOptions): DocEdit {
  return buildEdit(doc, (draft) => {
    const site = cellSite(doc, position);
    const table = site.table;
    const track = draft.track;
    if (track && !isOwnRow(track, table.rows[site.row])) {
      const struck = struckRow(track, table.rows[site.row]);
      if (struck !== table.rows[site.row]) {
        const marked = table.rows.slice();
        marked[site.row] = struck;
        draft.replace(site.story, site.container, site.index, 1, [tableWith(table, marked)]);
      }
      // The caret leaves the row that is deleted, as it would leave one that is gone.
      const to = site.row + 1 < table.rows.length ? site.row + 1 : site.row;
      return cellStart(doc, site, to, Math.min(site.cell, table.rows[to].cells.length - 1));
    }
    if (table.rows.length <= 1) return removeTable(draft, site);
    const rows = table.rows.slice();
    const gone = rows.splice(site.row, 1)[0];
    const harvest: Harvest = { pieces: [], blocks: [] };
    harvestRow(gone, harvest);
    const left = leftoverOf(harvest);
    const next = rows[site.row];
    if (next) {
      const spans = rowSpans(gone);
      const nextSpans = rowSpans(next);
      let cells: TableCell[] | undefined;
      for (let c = 0; c < gone.cells.length; c++) {
        if (gone.cells[c].props.vMerge !== 'restart') continue;
        for (let k = 0; k < next.cells.length; k++) {
          if (nextSpans[k][0] !== spans[c][0] || next.cells[k].props.vMerge !== 'continue') continue;
          if (!cells) cells = next.cells.slice();
          cells[k] = { ...next.cells[k], props: { ...next.cells[k].props, vMerge: 'restart' } };
        }
      }
      rows[site.row] = rowWith(cells ? { ...next, cells } : next, left, false);
    } else rows[site.row - 1] = rowWith(rows[site.row - 1], left, true);
    draft.replace(site.story, site.container, site.index, 1, [tableWith(table, rows)]);
    const row = Math.min(site.row, rows.length - 1);
    return cellStart(doc, site, row, Math.min(site.cell, rows[row].cells.length - 1));
  }, _options, true);
}

function widened(props: TableProps, by: number): TableProps {
  if (!props.width || props.width.type !== 'dxa') return props;
  return { ...props, width: { ...props.width, value: Math.max(0, props.width.value + by) } };
}

function cellWidened(props: CellProps, by: number): CellProps {
  if (!props.width || props.width.type !== 'dxa') return props;
  return { ...props, width: { ...props.width, value: Math.max(0, props.width.value + by) } };
}

/**
 * Inserts a column left or right of the cell a position is in, as wide as
 * the column beside it. Every row gets a cell at that place in the grid; a
 * cell that spans the place becomes one column wider instead, which is what
 * Word does with merged cells. The other columns keep their widths: the
 * table is one column wider, with or without tracking, also when it then is
 * wider than the text (UNSURE: not compared with Word, which may narrow the
 * columns of a table that fills the page instead).
 *
 * Where changes are tracked every new cell is a tracked insertion
 * (<w:cellIns>) and the table keeps the grid it had (<w:tblGridChange>). A
 * column that would widen a merged cell, or change the columns a row leaves
 * out, is refused there with TRACK_UNSUPPORTED: that is a change of the
 * properties of a cell or a row, which rejecting the new cells would not
 * take back.
 */
export function insertColumn(doc: DocModel, position: DocPosition, side: 'left' | 'right', _options?: DocEditOptions): DocEdit {
  if (side !== 'left' && side !== 'right') bad('A column goes left or right');
  return buildEdit(doc, (draft) => {
    const site = cellSite(doc, position);
    const table = site.table;
    if (table.grid.length >= MAX_COLUMNS) bad('The table has no room for another column');
    const refSpans = rowSpans(table.rows[site.row]);
    const [start, end] = refSpans[site.cell];
    const place = side === 'left' ? start : end;
    const beside = side === 'left' ? start : end - 1;
    const refWidth = table.rows[site.row].cells[site.cell].props.width;
    let width = table.grid[beside];
    if (!(width > 0)) width = refWidth && refWidth.type === 'dxa' && refWidth.value > 0 ? Math.round(refWidth.value / (end - start)) : 1440;
    const track = draft.track;
    const untracked = (): never => bad(TRACK_UNSUPPORTED);
    const rows = table.rows.map((row): TableRow => {
      const before = row.props.gridBefore || 0;
      if (place < before) return track ? untracked() : { ...row, props: { ...row.props, gridBefore: before + 1 } };
      const spans = rowSpans(row);
      const fresh = (model: TableCell): TableCell => {
        const props: CellProps = { ...plainProps(model.props) };
        delete props.gridSpan;
        delete props.hMerge;
        if (props.width) props.width = { type: 'dxa', value: width };
        const made: TableCell = { props, blocks: [emptyLike(model)] };
        return track ? insertedCell(track, made) : made;
      };
      for (let c = 0; c < row.cells.length; c++) {
        const [from, to] = spans[c];
        if (from === place) {
          const cells = row.cells.slice();
          cells.splice(c, 0, fresh(side === 'right' && c > 0 ? row.cells[c - 1] : row.cells[c]));
          return { ...row, cells };
        }
        if (from < place && place < to) {
          if (track) return untracked();
          const cells = row.cells.slice();
          cells[c] = { ...row.cells[c], props: { ...cellWidened(row.cells[c].props, width), gridSpan: to - from + 1 } };
          return { ...row, cells };
        }
      }
      const rowEnd = spans.length ? spans[spans.length - 1][1] : before;
      if (place === rowEnd && row.cells.length) return { ...row, cells: row.cells.concat([fresh(row.cells[row.cells.length - 1])]) };
      return track ? untracked() : { ...row, props: { ...row.props, gridAfter: (row.props.gridAfter || 0) + 1 } };
    });
    let grid = table.grid;
    if (place <= grid.length) {
      grid = grid.slice();
      grid.splice(place, 0, width);
    } else if (track) untracked();
    const next = tableWith(table, rows, grid, widened(table.props, width));
    // The grid as it was: kept once, for the first change of the columns (whoever made it). A grid the model keeps
    // as it was read (it holds more than columns) is written from that, and gets no record of the model.
    const keptGrid = (table.extras || []).some((extra) => extra.partial === true && extra.el.name.slice(extra.el.name.indexOf(':') + 1) === 'tblGrid');
    if (track && !table.gridChange && !keptGrid) next.gridChange = { id: track.revision().id, grid: table.grid.slice() };
    draft.replace(site.story, site.container, site.index, 1, [next]);
    return cellStart(doc, site, site.row, side === 'left' ? site.cell : site.cell + 1);
  }, _options, true);
}

/** Is the record of the grid of a table one this author made since the document was opened. */
const isOwnGrid = (track: Tracking, table: Table): boolean => table.gridChange !== undefined && track.made({ id: table.gridChange.id, author: track.author });

const sameGrid = (a: number[], b: number[]): boolean => a.length === b.length && a.every((width, i) => width === b[i]);

/**
 * deleteColumn() where changes are tracked: the cells of the columns stay,
 * as tracked deletions (<w:cellDel>) with everything in them deleted, and
 * the grid with them. Refused with TRACK_UNSUPPORTED where a cell covers
 * other columns too, or a row leaves some of the columns out, and where some
 * of the cells are deleted already and others are not (a column is one
 * change: it could not be resolved as one). Nothing happens to a column all
 * of whose cells are deleted already. Returns false when all the cells are
 * ones the author inserted in this session, which go for good.
 */
function strikeColumn(draft: Draft, track: Tracking, site: CellSite, start: number, end: number): boolean {
  const table = site.table;
  let own = true;
  const hit: [number, number][] = [];
  for (let r = 0; r < table.rows.length; r++) {
    const row = table.rows[r];
    const spans = rowSpans(row);
    const before = row.props.gridBefore || 0;
    const rowEnd = spans.length ? spans[spans.length - 1][1] : before;
    if (start < before || end > rowEnd) bad(TRACK_UNSUPPORTED);
    for (let c = 0; c < row.cells.length; c++) {
      const [from, to] = spans[c];
      if (to <= start || from >= end) continue;
      if (from < start || to > end) bad(TRACK_UNSUPPORTED);
      hit.push([r, c]);
      if (!isOwnCell(track, row.cells[c])) own = false;
    }
  }
  let struck = 0;
  for (const [r, c] of hit) if (table.rows[r].cells[c].props.deleted) struck++;
  if (struck === hit.length) return true;
  if (struck > 0) bad(TRACK_UNSUPPORTED);
  if (own) return false;
  const rows = table.rows.slice();
  let changed = false;
  for (const [r, c] of hit) {
    const cell = struckCell(track, rows[r].cells[c]);
    if (cell === rows[r].cells[c]) continue;
    const cells = rows[r].cells.slice();
    cells[c] = cell;
    rows[r] = { ...rows[r], cells };
    changed = true;
  }
  if (changed) draft.replace(site.story, site.container, site.index, 1, [tableWith(table, rows)]);
  return true;
}

/**
 * Deletes the grid columns the cell a position is in covers. Cells wholly in
 * those columns go; cells that also cover other columns become narrower. A
 * row left without cells goes, and the table with its last row. What the
 * removed cells held that is not content (marks of bookmarks and comments,
 * kept XML) moves into the cell after them in the row, else the one before
 * them, and of a row that goes into the row after or before it.
 */
export function deleteColumn(doc: DocModel, position: DocPosition, _options?: DocEditOptions): DocEdit {
  return buildEdit(doc, (draft) => {
    const site = cellSite(doc, position);
    const table = site.table;
    const [start, end] = rowSpans(table.rows[site.row])[site.cell];
    const track = draft.track;
    if (track && strikeColumn(draft, track, site, start, end)) {
      // The caret leaves the column that is deleted: to the cell after it, else the one before it.
      const cells = table.rows[site.row].cells.length;
      return cellStart(doc, site, site.row, site.cell + 1 < cells ? site.cell + 1 : Math.max(0, site.cell - 1));
    }
    const overlap = (from: number, to: number) => Math.max(0, Math.min(to, end) - Math.max(from, start));
    const widthOf = (from: number, to: number) => {
      let sum = 0;
      for (let g = Math.max(from, start); g < Math.min(to, end); g++) sum += table.grid[g] || 0;
      return sum;
    };
    const rows: TableRow[] = [];
    let caretRow = -1;
    /** What rows that went as a whole left behind, for the next row that stays. */
    let pending: Harvest = { pieces: [], blocks: [] };
    for (let r = 0; r < table.rows.length; r++) {
      const row = table.rows[r];
      const spans = rowSpans(row);
      const before = row.props.gridBefore || 0;
      const after = row.props.gridAfter || 0;
      const rowEnd = spans.length ? spans[spans.length - 1][1] : before;
      const cells: TableCell[] = [];
      const harvest: Harvest = { pieces: [], blocks: [] };
      /** Index among the cells that stay of the first one after the removed cells. */
      let behind = -1;
      for (let c = 0; c < row.cells.length; c++) {
        const [from, to] = spans[c];
        const cut = overlap(from, to);
        if (cut === 0) cells.push(row.cells[c]);
        else if (cut < to - from) {
          const props: CellProps = { ...cellWidened(row.cells[c].props, -widthOf(from, to)), gridSpan: to - from - cut };
          if (props.gridSpan === 1) delete props.gridSpan;
          cells.push({ ...row.cells[c], props });
        } else {
          harvestBlocks(row.cells[c].blocks, harvest);
          if (behind < 0) behind = cells.length;
        }
      }
      if (cells.length === 0) {
        harvestRow(row, pending);
        continue;
      }
      if (behind >= 0) {
        const left = leftoverOf(harvest);
        if (behind < cells.length) cells[behind] = cellWith(cells[behind], left, false);
        else cells[cells.length - 1] = cellWith(cells[cells.length - 1], left, true);
      }
      const cutBefore = overlap(0, before);
      const cutAfter = overlap(rowEnd, rowEnd + after);
      let props = row.props;
      if (cutBefore || cutAfter) {
        props = { ...props };
        if (cutBefore) props.gridBefore = before - cutBefore;
        if (cutAfter) props.gridAfter = after - cutAfter;
      }
      if (r === site.row) caretRow = rows.length;
      let kept: TableRow = cells.length === row.cells.length && props === row.props && cells.every((cell, i) => cell === row.cells[i]) ? row : { ...row, props, cells };
      if (pending.pieces.length || pending.blocks.length) {
        kept = rowWith(kept, leftoverOf(pending), false);
        pending = { pieces: [], blocks: [] };
      }
      rows.push(kept);
    }
    if (rows.length === 0) return removeTable(draft, site);
    if (pending.pieces.length || pending.blocks.length) rows[rows.length - 1] = rowWith(rows[rows.length - 1], leftoverOf(pending), true);
    const grid = table.grid.slice();
    if (start < grid.length) grid.splice(start, end - start);
    const next = tableWith(table, rows, grid, widened(table.props, -widthOf(start, end)));
    // The author took back the columns they had put in: the grid is what their record says, and the record goes.
    if (track && isOwnGrid(track, table) && sameGrid(grid, (table.gridChange as { grid: number[] }).grid)) delete next.gridChange;
    draft.replace(site.story, site.container, site.index, 1, [next]);
    const row = caretRow >= 0 ? caretRow : Math.min(site.row, rows.length - 1);
    return cellStart(doc, site, row, Math.min(site.cell, rows[row].cells.length - 1));
  }, _options, true);
}

const SINGLE: Border = Object.freeze({ style: 'single', size: 4, space: 0, color: Object.freeze({ auto: true }) });

/** Width of the text area at a position, twips: of the cell it is in, else of the page between its margins. */
function availableWidth(doc: DocModel, position: DocPosition): number {
  const path = position.path;
  if (path.length >= 4) {
    const table = containerBlocks(doc, position.story, path.slice(0, -4))[path[path.length - 4]];
    if (table && table.type === 'table') {
      const row = table.rows[path[path.length - 3]];
      const cell = row.cells[path[path.length - 2]];
      const [from, to] = rowSpans(row)[path[path.length - 2]];
      let width = cell.props.width && cell.props.width.type === 'dxa' ? cell.props.width.value : 0;
      if (!(width > 0)) for (let g = from; g < to; g++) width += table.grid[g] || 0;
      // Less the default cell margins of Word on both sides.
      if (width > 0) return Math.max(360, width - 216);
    }
  }
  const section = doc.body.section;
  const page = (section && section.page && section.page.width) || 12240;
  const margins = (section && section.margins) || {};
  const left = margins.left === undefined ? 1440 : margins.left;
  const right = margins.right === undefined ? 1440 : margins.right;
  return Math.max(360, page - left - right);
}

/**
 * Inserts a table of `rows` by `columns` empty cells of equal width at a
 * position (see insertBlock for where exactly). The table uses the table
 * style "TableGrid" when the document has it and single borders of its own
 * otherwise. The caret goes into the first cell. Where changes are tracked
 * the table is a tracked insertion: every row of it says so, and so does the
 * mark of every paragraph the command makes.
 */
export function insertTable(doc: DocModel, position: DocPosition, rows: number, columns: number, _options?: DocEditOptions): DocEdit {
  if (!whole(rows, 1, MAX_NEW_CELLS) || !whole(columns, 1, MAX_COLUMNS) || rows * columns > MAX_NEW_CELLS) bad('The size of the table is not valid');
  return buildEdit(doc, (draft) => {
    resolvePosition(doc, position);
    const width = Math.max(1, Math.floor(availableWidth(doc, position) / columns));
    const style = doc.styles.styles.get('TableGrid');
    const props: TableProps = { width: { type: 'auto', value: 0 } };
    if (style && style.type === 'table') {
      props.style = 'TableGrid';
      props.look = { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true, val: 0x04a0 };
    } else props.borders = { top: SINGLE, left: SINGLE, bottom: SINGLE, right: SINGLE, insideH: SINGLE, insideV: SINGLE };
    const grid: number[] = [];
    for (let c = 0; c < columns; c++) grid.push(width);
    const cellProps: CellProps = { width: { type: 'dxa', value: width } };
    const list: TableRow[] = [];
    for (let r = 0; r < rows; r++) {
      const cells: TableCell[] = [];
      for (let c = 0; c < columns; c++) cells.push({ props: cellProps, blocks: [{ type: 'paragraph', props: EMPTY_PARA_PROPS, children: [] }] });
      list.push({ props: {}, cells });
    }
    const table: Table = { type: 'table', props, grid, rows: list };
    const index = placeBlock(draft, position, draft.track ? insertedTable(draft.track, table) : table);
    return { story: position.story, path: position.path.slice(0, -1).concat([index, 0, 0, 0]), offset: 0 };
  }, _options, true);
}

// ------------------------------------------------------------------ hyperlinks

const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_BASE_STRICT = 'http://purl.oclc.org/ooxml/officeDocument/relationships';

/** An id no relationship of the part has yet. */
export function freeRelId(relationships: Map<string, unknown>): string {
  let next = 1;
  relationships.forEach((_value, id) => {
    const match = /^rId(\d{1,9})$/.exec(id);
    if (match && Number(match[1]) >= next) next = Number(match[1]) + 1;
  });
  return 'rId' + next;
}

const isRevision = (w: Wrapper) => w.kind === 'ins' || w.kind === 'del' || w.kind === 'moveFrom' || w.kind === 'moveTo';

/** Items a tracked link takes a second copy of: text and what is typed like it. */
const LINKABLE = new Set<string>(['text', 'tab', 'break', 'noBreakHyphen', 'softHyphen', 'symbol']);

/** Is content the result of a simple field, or inside a content control: what a second copy of its text beside it would write twice, or cut. */
function inFieldOrControl(wrap: Wrapper | undefined): boolean {
  let steps = 0;
  for (let w = wrap; w && steps < 4096; w = w.parent, steps++) if (w.kind === 'simpleField' || w.kind === 'sdt' || w.kind === 'customXml') return true;
  return false;
}

/** Would taking content out of the tracked changes around it cut a content control in two (it sits inside one of them). */
function cutsControl(wrap: Wrapper | undefined): boolean {
  let control = false;
  let steps = 0;
  for (let w = wrap; w && steps < 4096; w = w.parent, steps++) {
    if (w.kind === 'sdt' || w.kind === 'customXml') control = true;
    else if (control && isRevision(w)) return true;
  }
  return false;
}

/**
 * Makes the text of a range a hyperlink to `target`: an address ("https://…",
 * "mailto:…", "tel:…"), or "#name" for a bookmark of the document. The range
 * must lie in one paragraph; a collapsed range gets the target itself as its
 * text. A link the text already has is replaced. For an address the part of
 * the story gets a relationship, which undoing removes again.
 *
 * Only a target that is safe to follow is taken (safeHyperlinkTarget: http,
 * https, mailto, tel, or a bookmark): a script or a file address throws
 * RangeError, so that no document leaves the editor with one. The target is
 * stored as that function reads it, without the space around it.
 *
 * Where changes are tracked: the element of a link carries no revision, so a
 * link over text that is there is that text deleted and the same text
 * inserted inside the link, which is how Word tracks it (UNSURE: seen in
 * files, not compared with Word here). Rejecting the two gives the text back
 * as it was; another address for a link is the same, with the deleted text
 * left in the old link. Text the editor typed in this session leaves no
 * deletion behind. A range that holds more than text (a picture, a note
 * reference, an equation, text of a field, simple or complex), and one that
 * holds text of a content control, is refused there with TRACK_UNSUPPORTED:
 * a second copy of those would be a second picture, a second reference,
 * half a field, a second field or control.
 */
export function insertHyperlink(doc: DocModel, range: DocRange, address: string, _options?: DocEditOptions): DocEdit {
  let plain = typeof address === 'string' && address.length <= 2083;
  // Control characters have no place in an address; a browser would drop some of them and read another address.
  for (let i = 0; plain && i < address.length; i++) if (address.charCodeAt(i) < 0x20) plain = false;
  const safe = plain ? safeHyperlinkTarget(address) : undefined;
  if (safe === undefined || safe === '' || safe === '#') return bad('The target of the hyperlink is not valid');
  const target: string = safe;
  const at = asRange(range);
  return buildEdit(doc, (draft) => {
    const { start, end } = orderRange(at);
    const where = resolvePosition(doc, start);
    resolvePosition(doc, end);
    if (inTextBox(start.path)) bad('The content of a text box cannot be edited');
    if (start.path.length !== end.path.length || start.path.some((value, i) => value !== end.path[i])) bad('A hyperlink must lie in one paragraph');
    const story = start.story;
    let link: Wrapper;
    if (target.charAt(0) === '#') link = { kind: 'hyperlink', anchor: target.slice(1) };
    else {
      const host = storyOf(doc, story).host;
      const id = freeRelId(host.relationships);
      draft.apply({ type: 'setRelationship', story, id, relationship: { id, type: (doc.strict ? REL_BASE_STRICT : REL_BASE) + '/hyperlink', target, external: true, part: null } });
      link = { kind: 'hyperlink', rId: id, target };
    }
    const paragraph = where.paragraph;
    const pieces = flatten(paragraph);
    const caret = start.offset === end.offset && (start.before === true) === (end.before === true);
    const from = caret ? caretIndex(pieces, start) : startIndex(pieces, start);
    const styled = doc.styles.styles.get('Hyperlink');
    const useStyle = styled !== undefined && styled.type === 'character';
    const restyled = (props: RunProps): RunProps => (useStyle ? patchRunProps(props, { style: 'Hyperlink' }) : patchRunProps(props, { color: '0563C1', underline: true }));
    const shown = target.replace(/[\t\n\u000b]/g, ' ');
    let to: number;
    if (caret && !draft.track) {
      const props = inheritedProps(pieces, from, paragraph.props.markProps);
      const wrap = wrapBetween(from > 0 ? pieces[from - 1].wrap : undefined, from < pieces.length ? pieces[from].wrap : undefined);
      const added = textPieces(shown, props, wrap, {});
      pieces.splice(from, 0, ...added);
      to = from + added.length;
    } else to = caret ? from : caretIndex(pieces, end);
    // The link goes inside the wrappers the text is in, but around tracked changes (which may only hold runs), and
    // replaces a link that is there. Wrappers whose parent changes are copied, once each, so that pieces that
    // shared a wrapper still do.
    const stripped = new Map<Wrapper, Wrapper | undefined>();
    const strip = (w: Wrapper | undefined): Wrapper | undefined => {
      if (!w) return undefined;
      if (stripped.has(w)) return stripped.get(w);
      const parent = strip(w.parent);
      let out: Wrapper | undefined;
      if (isRevision(w) || w.kind === 'hyperlink') out = parent;
      else if (parent === w.parent) out = w;
      else {
        out = { ...w };
        if (parent) out.parent = parent;
        else delete out.parent;
      }
      stripped.set(w, out);
      return out;
    };
    const links = new Map<Wrapper | undefined, Wrapper>();
    const linkIn = (outer: Wrapper | undefined): Wrapper => {
      let mine = links.get(outer);
      if (!mine) {
        mine = { ...link };
        if (outer) mine.parent = outer;
        links.set(outer, mine);
      }
      return mine;
    };
    const track = draft.track;
    if (track) {
      // The linked text is new text: each stretch of it in an insertion of its own, inside the link.
      const insertions = new Map<Wrapper, Wrapper>();
      const insertionIn = (outer: Wrapper | undefined): Wrapper => {
        if (cutsControl(outer)) bad(TRACK_UNSUPPORTED);
        const inLink = linkIn(strip(outer));
        let mine = insertions.get(inLink);
        if (!mine) insertions.set(inLink, (mine = { kind: 'ins', ...track.revision(), parent: inLink }));
        return mine;
      };
      let added: Piece[] = [];
      let list = pieces;
      let at = from;
      if (caret) {
        const wrap = wrapBetween(from > 0 ? pieces[from - 1].wrap : undefined, from < pieces.length ? pieces[from].wrap : undefined);
        added = textPieces(shown, restyled(inheritedProps(pieces, from, paragraph.props.markProps)), insertionIn(wrap), {});
      } else {
        const owner = fieldPairs(pieces).owner;
        const styles = new Map<RunProps, RunProps>();
        const groups = new Map<object, object>();
        for (let q = from; q < to; q++) {
          const piece = pieces[q];
          // What is no content (a bookmark, deleted text) stays where it is, outside the link.
          if (piece.frozen || !isRemovable(piece)) continue;
          const item = piece.item;
          if (!item || owner[q] >= 0 || !LINKABLE.has(item.type) || inFieldOrControl(piece.wrap)) return bad(TRACK_UNSUPPORTED);
          let props = styles.get(piece.props);
          if (!props) styles.set(piece.props, (props = restyled(piece.props)));
          let group = groups.get(piece.group);
          if (!group) groups.set(piece.group, (group = {}));
          added.push({ item, props, wrap: insertionIn(piece.wrap), group, index: -1, code: false, len: piece.len });
        }
        const struck = strikePieces(track, pieces, from, to);
        list = struck.pieces;
        // The linked text follows what was deleted: before the piece that followed the range, wherever that is now.
        at = to < pieces.length ? list.indexOf(pieces[to]) : list.length;
        if (at < 0) at = list.length - (pieces.length - to - 1);
      }
      list = list.slice();
      list.splice(at, 0, ...added);
      const first = offsetOf(list, at);
      const last = offsetOf(list, at + added.length);
      draft.setParagraph(story, start.path, paragraphWith(paragraph, rebuild(list)));
      return { anchor: { story, path: start.path.slice(), offset: first }, focus: { story, path: start.path.slice(), offset: last } };
    }
    const revisions = new Map<Wrapper, Wrapper>();
    const linked = (w: Wrapper | undefined, base: Wrapper): Wrapper => {
      if (!w) return base;
      if (!isRevision(w)) return linked(w.parent, base);
      let copy = revisions.get(w);
      if (!copy) revisions.set(w, (copy = { ...w, parent: linked(w.parent, base) }));
      return copy;
    };
    for (let q = from; q < to; q++) {
      const piece = pieces[q];
      if (piece.frozen) continue;
      const next: Piece = { ...piece, wrap: linked(piece.wrap, linkIn(strip(piece.wrap))) };
      if (piece.item && !piece.code && piece.item.type !== 'fieldChar') next.props = restyled(piece.props);
      pieces[q] = next;
    }
    const first = offsetOf(pieces, from);
    const last = offsetOf(pieces, to);
    draft.setParagraph(story, start.path, paragraphWith(paragraph, rebuild(pieces)));
    return { anchor: { story, path: start.path.slice(), offset: first }, focus: { story, path: start.path.slice(), offset: last } };
  }, _options, true);
}

// -------------------------------------------------------------------- pictures

export interface NewImage {
  bytes: Uint8Array;
  /** Media type: "image/png", "image/jpeg", "image/gif", "image/bmp" or "image/webp". The bytes have to begin the way a file of that type does. */
  type: string;
  /** Size as shown, in EMU (914400 to the inch, 9525 to the CSS pixel): whole numbers from 1 to 91440000 (100 inches). */
  width: number;
  height: number;
  /** Alternative text. */
  description?: string;
}

export interface InsertImageOptions extends DocEditOptions {
  /**
   * Is a part of this name ("/word/media/image3.png") in the package already?
   * The commands see the model, not the package: a caller that has the
   * package says so here, and the new picture gets a name that is free.
   */
  taken?: (part: string) => boolean;
}

const startsWith = (bytes: Uint8Array, signature: number[], at = 0): boolean => {
  if (bytes.length < at + signature.length) return false;
  for (let i = 0; i < signature.length; i++) if (bytes[at + i] !== signature[i]) return false;
  return true;
};

/** The picture formats every browser shows, with the extension of their part and how a file of the type begins. */
const IMAGE_TYPES = new Map<string, { extension: string; is: (bytes: Uint8Array) => boolean }>([
  ['image/png', { extension: 'png', is: (bytes) => startsWith(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]) }],
  ['image/jpeg', { extension: 'jpeg', is: (bytes) => startsWith(bytes, [0xff, 0xd8, 0xff]) }],
  ['image/gif', { extension: 'gif', is: (bytes) => startsWith(bytes, [0x47, 0x49, 0x46, 0x38]) && (bytes[4] === 0x37 || bytes[4] === 0x39) && bytes[5] === 0x61 }],
  ['image/bmp', { extension: 'bmp', is: (bytes) => startsWith(bytes, [0x42, 0x4d]) && bytes.length >= 6 }],
  ['image/webp', { extension: 'webp', is: (bytes) => startsWith(bytes, [0x52, 0x49, 0x46, 0x46]) && startsWith(bytes, [0x57, 0x45, 0x42, 0x50], 8) }],
]);

/** Largest side of a picture, EMU: 100 inches, several times the largest page of Word. */
const MAX_IMAGE_EMU = 91440000;
const MAX_DESCRIPTION = 4096;

const DRAWING_NS = {
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  r: REL_BASE,
};
const DRAWING_NS_STRICT = {
  wp: 'http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing',
  a: 'http://purl.oclc.org/ooxml/drawingml/main',
  pic: 'http://purl.oclc.org/ooxml/drawingml/picture',
  r: REL_BASE_STRICT,
};

const el = (name: string, attrs: [string, string][], children: XmlNode[] = []): XmlElement => ({ type: 'element', name, attrs: new Map(attrs), children });

/** What the document has in use that a new picture must not take: names of media parts and ids of drawings. */
export interface Taken {
  /** Highest N among the parts named ".../media/imageN.*". */
  image: number;
  /** Names of parts the relationships and DocModel.media refer to, in lower case. */
  parts: Set<string>;
  /** Highest id of a drawing or of an object in one (wp:docPr, pic:cNvPr, ...). */
  drawing: number;
}

/**
 * Walks every story of the document (the comments too), into text boxes and
 * into the trees kept of drawings, legacy pictures and alternate content,
 * for the names and ids in use.
 */
export function takenByDocument(doc: DocModel): Taken {
  const taken: Taken = { image: 0, parts: new Set<string>(), drawing: 0 };
  const part = (name: string | null | undefined) => {
    if (!name) return;
    taken.parts.add(name.toLowerCase());
    const match = /\/media\/image(\d{1,9})\.[^/]*$/i.exec(name);
    if (match && Number(match[1]) > taken.image) taken.image = Number(match[1]);
  };
  if (doc.media) doc.media.forEach((_media, name) => part(name));
  const hosts: { relationships: Map<string, { part: string | null }> }[] = [doc.body];
  doc.headers.forEach((story) => hosts.push(story));
  doc.footers.forEach((story) => hosts.push(story));
  for (const collection of [doc.footnotes, doc.endnotes, doc.comments]) if (collection) hosts.push(collection);
  for (const host of hosts) host.relationships.forEach((relationship) => part(relationship.part));

  /** Every id of a docPr or cNvPr element in a kept tree, however deep. */
  const tree = (root: XmlElement) => {
    const stack: XmlElement[] = [root];
    while (stack.length) {
      const node = stack.pop() as XmlElement;
      const local = node.name.slice(node.name.indexOf(':') + 1);
      if (local === 'docPr' || local === 'cNvPr') {
        const text = node.attrs.get('id');
        const id = text !== undefined && /^\d{1,10}$/.test(text) ? Number(text) : 0;
        if (id > taken.drawing && id < 0x7fffffff) taken.drawing = id;
      }
      for (const child of node.children) if (child.type === 'element') stack.push(child);
    }
  };
  const wrappers = new Set<Wrapper>();
  const wrapped = (wrap: Wrapper | undefined) => {
    for (let w = wrap; w && !wrappers.has(w); w = w.parent) {
      wrappers.add(w);
      if (w.kind === 'alternate') tree(w.xml);
    }
  };
  const items = (list: RunItem[], depth: number) => {
    for (const item of list) {
      if (item.type === 'alternate') {
        tree(item.xml);
        if (depth < 64) items(item.items, depth + 1);
      } else if (item.type === 'drawing' || item.type === 'pict' || item.type === 'object') {
        tree(item.xml);
        if (item.textBoxes && depth < 64) for (const box of item.textBoxes) blocks(box, depth + 1);
      } else if (item.type === 'unknown') tree(item.xml);
    }
  };
  const blocks = (list: Block[], depth: number) => {
    if (depth > 256) return;
    for (const block of list) {
      wrapped(block.wrap);
      if (block.type === 'paragraph') {
        for (const inline of block.children) {
          wrapped(inline.wrap);
          if (inline.type === 'run') items(inline.content, depth);
          else if (inline.type === 'unknown') tree(inline.xml);
        }
      } else if (block.type === 'table') {
        for (const row of block.rows) for (const cell of row.cells) blocks(cell.blocks, depth + 1);
      } else if (block.type === 'unknown') tree(block.xml);
    }
  };
  blocks(doc.body.blocks, 0);
  const stories = new Set<{ blocks: Block[] }>();
  doc.headers.forEach((story) => stories.add(story));
  doc.footers.forEach((story) => stories.add(story));
  stories.forEach((story) => blocks(story.blocks, 0));
  for (const collection of [doc.footnotes, doc.endnotes, doc.comments]) if (collection) collection.items.forEach((item: { blocks: Block[] }) => blocks(item.blocks, 0));
  return taken;
}

/**
 * Inserts a picture in line with the text at a position.
 *
 * The bytes go into DocModel.media under the name of a new package part,
 * "/word/media/imageN.ext", and the part of the story gets a relationship to
 * it; the drawing refers to both (ImageRef.rId, ImageRef.target). A writer
 * has to add every entry of DocModel.media that the package does not hold as
 * a part, with a content type for its extension. Undoing removes the entry
 * and the relationship again.
 *
 * N goes on from the highest number among the media parts the model knows
 * of: those of DocModel.media and of the relationships of every story. Parts
 * of the package that no story refers to (a picture bullet of the numbering
 * part) are not in the model: `options.taken` lets the caller answer for the
 * package, and a name it reports as taken is passed over.
 *
 * Throws RangeError for a type other than png, jpeg, gif, bmp and webp, for
 * bytes that do not begin like a file of the type, and for a size that is
 * not a whole number of EMU between 1 and 91440000.
 *
 * Where changes are tracked the picture sits in a tracked insertion.
 * Rejecting that takes the picture out of its paragraph; the entry of
 * DocModel.media and the relationship stay, as they do when a picture is
 * deleted (undoing needs them), and a writer writes a part nothing refers to.
 */
export function insertImage(doc: DocModel, position: DocPosition, image: NewImage, options?: InsertImageOptions): DocEdit {
  if (!image || typeof image !== 'object') bad('A picture needs its bytes');
  if (!(image.bytes instanceof Uint8Array) || image.bytes.length === 0) bad('A picture needs its bytes');
  const kind = typeof image.type === 'string' ? IMAGE_TYPES.get(image.type) : undefined;
  if (!kind) return bad('The type of the picture is not supported');
  if (!kind.is(image.bytes)) bad('The bytes of the picture are not of its type');
  if (!whole(image.width, 1, MAX_IMAGE_EMU) || !whole(image.height, 1, MAX_IMAGE_EMU)) bad('The size of the picture is not valid');
  if (image.description !== undefined && (typeof image.description !== 'string' || image.description.length > MAX_DESCRIPTION)) bad('The description of the picture is not valid');
  const isTaken = options ? options.taken : undefined;
  if (isTaken !== undefined && typeof isTaken !== 'function') bad('The check for names in use must be a function');
  const description = image.description === undefined ? undefined : cleanText(image.description);
  return buildEdit(doc, (draft) => {
    const at = resolvePosition(doc, position);
    if (inTextBox(position.path)) bad('The content of a text box cannot be edited');
    const story = position.story;
    const host = storyOf(doc, story).host;
    const taken = takenByDocument(doc);
    if (taken.drawing >= 0x7ffffffe) bad('The document has no room for another drawing');
    let number = taken.image + 1;
    let name = `/word/media/image${number}.${kind.extension}`;
    // Each turn tries the next number; the caller is asked a bounded number of times.
    for (let tries = 0; taken.parts.has(name.toLowerCase()) || (isTaken !== undefined && isTaken(name) === true); tries++) {
      if (tries >= 10000) bad('No name is free for the picture');
      number++;
      name = `/word/media/image${number}.${kind.extension}`;
    }
    const media: DocMedia = { bytes: image.bytes, type: image.type };
    draft.apply({ type: 'setMedia', name, media });
    const rId = freeRelId(host.relationships);
    // Relative to the part of the story when that is beside the media folder, as Word writes it.
    const beside = /^\/word\/[^/]+$/.test(host.part);
    draft.apply({
      type: 'setRelationship',
      story,
      id: rId,
      relationship: { id: rId, type: (doc.strict ? REL_BASE_STRICT : REL_BASE) + '/image', target: beside ? name.slice('/word/'.length) : name, external: false, part: name },
    });

    const ns = doc.strict ? DRAWING_NS_STRICT : DRAWING_NS;
    let w = 'w';
    for (const [attr, value] of host.rootAttrs) if (attr.indexOf('xmlns:') === 0 && /wordprocessingml\/(2006\/)?main$/.test(value)) w = attr.slice(6);
    const id = String(taken.drawing + 1);
    const title = `Picture ${id}`;
    const size: [string, string][] = [
      ['cx', String(image.width)],
      ['cy', String(image.height)],
    ];
    const docPr: [string, string][] = [
      ['id', id],
      ['name', title],
    ];
    if (description !== undefined) docPr.push(['descr', description]);
    const xml = el(`${w}:drawing`, [], [
      el('wp:inline', [['xmlns:wp', ns.wp], ['distT', '0'], ['distB', '0'], ['distL', '0'], ['distR', '0']], [
        el('wp:extent', size),
        el('wp:docPr', docPr),
        el('wp:cNvGraphicFramePr', [], [el('a:graphicFrameLocks', [['xmlns:a', ns.a], ['noChangeAspect', '1']])]),
        el('a:graphic', [['xmlns:a', ns.a]], [
          el('a:graphicData', [['uri', ns.pic]], [
            el('pic:pic', [['xmlns:pic', ns.pic]], [
              el('pic:nvPicPr', [], [el('pic:cNvPr', [['id', id], ['name', title]]), el('pic:cNvPicPr', [])]),
              el('pic:blipFill', [], [el('a:blip', [['xmlns:r', ns.r], ['r:embed', rId]]), el('a:stretch', [], [el('a:fillRect', [])])]),
              el('pic:spPr', [], [
                el('a:xfrm', [], [el('a:off', [['x', '0'], ['y', '0']]), el('a:ext', size)]),
                el('a:prstGeom', [['prst', 'rect']], [el('a:avLst', [])]),
              ]),
            ]),
          ]),
        ]),
      ]),
    ]);
    const ref = { rId, target: name };
    const item: DrawingItem = { type: 'drawing', placement: 'inline', kind: 'picture', width: image.width, height: image.height, name: title, images: [{ image: ref }], image: ref, xml };
    if (description !== undefined) item.description = description;

    const paragraph = at.paragraph;
    let pieces = flatten(paragraph);
    const index = caretIndex(pieces, position);
    const props = inheritedProps(pieces, index, paragraph.props.markProps);
    let wrap = wrapBetween(index > 0 ? pieces[index - 1].wrap : undefined, index < pieces.length ? pieces[index].wrap : undefined);
    if (draft.track) {
      // The picture is a tracked insertion, like text typed at the same place.
      const placed = insertionWrap(draft.track, pieces, index, wrap);
      pieces = placed.pieces.slice();
      wrap = placed.wrap;
    }
    pieces.splice(index, 0, itemPiece(item, props, wrap, groupFor(pieces, index, props, wrap)));
    draft.setParagraph(story, position.path, paragraphWith(paragraph, rebuild(pieces)));
    return { story, path: position.path.slice(), offset: offsetOf(pieces, index) };
  }, options, true);
}
