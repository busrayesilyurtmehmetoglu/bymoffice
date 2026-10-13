// Tracked changes of whole blocks: paragraphs, tables, rows and cells that
// are put in or taken out as revisions (see track.ts for text).
//
// What such a change looks like in the model:
//   - a row that was put in says `inserted` in its properties (<w:ins> in
//     <w:trPr>), and so does the mark of every paragraph in its cells; a
//     table that was put in is a table all of whose rows say so;
//   - a row that was taken out stays and says `deleted` (<w:del> in <w:trPr>);
//     what its cells hold is deleted like any other content: the text in
//     deletions, the paragraph marks deleted. UNSURE: this is what files
//     written by Word show; it was not compared with Word here;
//   - a cell that was put in or taken out with its column says `inserted` or
//     `deleted` (<w:cellIns>, <w:cellDel> in <w:tcPr>); the content of a cell
//     that was taken out is deleted as in a row. The grid of the table as it
//     was before a column went in is kept (Table.gridChange, <w:tblGridChange>);
//   - content that is put in whole (a paste, a picture, a linked text) sits in
//     insertions, innermost around its runs, since an insertion holds nothing
//     but runs.
// A row, and a table, the same author put in since the document was opened is
// taken out for good when nothing in it is anybody else's.

import { EMPTY_PARA_PROPS, type Block, type CellProps, type Inline, type ParaProps, type Paragraph, type RowProps, type Table, type TableCell, type TableRow, type Wrapper } from '../../model/doc/types';
import { READ_ONLY, flatten, paragraphWith, rebuild, type Piece } from './pieces';
import { isDeletedMark, isOwnParagraph, strikePieces, withMarkRevision, type Tracking } from './track';

const MAX_DEPTH = 64;

const bad = (message: string): never => {
  throw new RangeError(message);
};

function frozen(wrap: Wrapper | undefined): boolean {
  let steps = 0;
  for (let w = wrap; w && steps < 4096; w = w.parent, steps++) if (w.kind === 'alternate') return true;
  return false;
}

// ------------------------------------------------------------------ deletions

/** A paragraph with everything in it deleted as a tracked change, and (with `mark`) its mark too. The paragraph itself when nothing changes. */
function struckParagraph(track: Tracking, paragraph: Paragraph, mark: boolean): Paragraph {
  if (frozen(paragraph.wrap)) return bad(READ_ONLY);
  const pieces = flatten(paragraph);
  const struck = strikePieces(track, pieces, 0, pieces.length);
  let props = paragraph.props;
  if (mark && !isDeletedMark(props)) props = withMarkRevision(props, 'deleted', track.revision());
  if (!struck.changed && props === paragraph.props) return paragraph;
  return paragraphWith(paragraph, struck.changed ? rebuild(struck.pieces) : paragraph.children, props);
}

/**
 * Blocks with everything in them deleted as a tracked change: the content of
 * their paragraphs, the marks of the paragraphs (`lastMark`: also the mark of
 * the last one, which a cell that stays keeps), the rows of their tables. The
 * list itself when nothing changes. Throws RangeError(READ_ONLY) for content
 * that is written back whole.
 */
export function struckBlocks(track: Tracking, blocks: Block[], lastMark: boolean, depth = 0): Block[] {
  if (depth > MAX_DEPTH) return bad('The tables are nested too deep');
  let last = -1;
  if (!lastMark) for (let i = blocks.length - 1; i >= 0 && last < 0; i--) if (blocks[i].type === 'paragraph') last = i;
  let out: Block[] | undefined;
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    let next: Block = block;
    if (block.type === 'paragraph') next = struckParagraph(track, block, i !== last);
    else if (block.type === 'table') next = struckTable(track, block, depth + 1);
    if (next === block) continue;
    if (!out) out = blocks.slice();
    out[i] = next;
  }
  return out || blocks;
}

/**
 * A row that says it was deleted, with everything in its cells deleted. A row that is deleted
 * already, by whomever, is left as it is, with what it holds: nothing is deleted twice.
 */
export function struckRow(track: Tracking, row: TableRow, depth = 0): TableRow {
  if (frozen(row.wrap)) return bad(READ_ONLY);
  if (row.props.deleted) return row;
  let cells: TableCell[] | undefined;
  for (let c = 0; c < row.cells.length; c++) {
    const cell = row.cells[c];
    if (frozen(cell.wrap)) return bad(READ_ONLY);
    const blocks = struckBlocks(track, cell.blocks, true, depth);
    if (blocks === cell.blocks) continue;
    if (!cells) cells = row.cells.slice();
    cells[c] = { ...cell, blocks };
  }
  const props: RowProps = { ...row.props, deleted: track.revision() };
  return { ...row, props, cells: cells || row.cells };
}

/** A table all of whose rows say they were deleted. The table itself when it is like that already. */
export function struckTable(track: Tracking, table: Table, depth = 0): Table {
  if (frozen(table.wrap)) return bad(READ_ONLY);
  let rows: TableRow[] | undefined;
  for (let r = 0; r < table.rows.length; r++) {
    const row = struckRow(track, table.rows[r], depth);
    if (row === table.rows[r]) continue;
    if (!rows) rows = table.rows.slice();
    rows[r] = row;
  }
  if (!rows) return table;
  const out: Table = { ...table, rows };
  delete out.srcStart;
  delete out.srcEnd;
  return out;
}

/** A cell that says it was deleted (with its column), with everything in it deleted. (One that is deleted already does not come here: see strikeColumn.) */
export function struckCell(track: Tracking, cell: TableCell, depth = 0): TableCell {
  if (frozen(cell.wrap)) return bad(READ_ONLY);
  const blocks = struckBlocks(track, cell.blocks, true, depth);
  const props: CellProps = { ...cell.props, deleted: track.revision() };
  return { ...cell, props, blocks };
}

// ----------------------------------------------------------------- insertions

/** Paragraph properties whose mark says that this command inserted it. */
function withInsertedMark(track: Tracking, props: ParaProps): ParaProps {
  return withMarkRevision(props, 'inserted', track.revision());
}

/**
 * Pieces as content this command puts in: each in an insertion, innermost
 * around it (an insertion holds runs and nothing else, so it goes inside a
 * link or a control the piece is in). Pieces that follow each other in one
 * wrapper share one insertion. What is no content (the marks of ranges) and
 * what is written back whole stays as it is.
 */
export function insertedPieces(track: Tracking, pieces: Piece[]): Piece[] {
  const out: Piece[] = [];
  let open: { parent: Wrapper | undefined; wrap: Wrapper } | undefined;
  for (const piece of pieces) {
    const content = piece.item !== undefined || (piece.inline !== undefined && (piece.inline.type === 'run' || piece.inline.type === 'math'));
    if (!content || piece.frozen || frozen(piece.wrap)) {
      out.push(piece);
      open = undefined;
      continue;
    }
    if (!open || open.parent !== piece.wrap) {
      const wrap = { kind: 'ins', ...track.revision() } as Wrapper;
      if (piece.wrap) wrap.parent = piece.wrap;
      open = { parent: piece.wrap, wrap };
    }
    const next: Piece = { ...piece, wrap: open.wrap };
    if (piece.inline) {
      const inline = { ...piece.inline, wrap: open.wrap } as Inline;
      next.inline = inline;
      next.group = inline;
    }
    out.push(next);
  }
  return out;
}

/** A paragraph this command puts in whole: its content in insertions and (with `mark`) its mark inserted. */
export function insertedParagraph(track: Tracking, paragraph: Paragraph, mark: boolean): Paragraph {
  const pieces = flatten(paragraph);
  const children = pieces.length ? rebuild(insertedPieces(track, pieces)) : paragraph.children;
  const props = mark ? withInsertedMark(track, paragraph.props) : paragraph.props;
  return paragraphWith(paragraph, children, props);
}

/** Blocks this command puts in whole: every paragraph and its mark, every row of every table. */
export function insertedBlocks(track: Tracking, blocks: Block[], depth = 0): Block[] {
  if (depth > MAX_DEPTH) return bad('The tables are nested too deep');
  return blocks.map((block): Block => {
    if (block.type === 'paragraph') return insertedParagraph(track, block, true);
    if (block.type === 'table') return insertedTable(track, block, depth + 1);
    return block;
  });
}

/** A row this command puts in: it says so, and so does everything in its cells. */
export function insertedRow(track: Tracking, row: TableRow, depth = 0): TableRow {
  const props: RowProps = { ...row.props, inserted: track.revision() };
  delete props.deleted;
  return { ...row, props, cells: row.cells.map((cell) => ({ ...cell, blocks: insertedBlocks(track, cell.blocks, depth) })) };
}

/** A table this command puts in: every row of it is an inserted row. */
export function insertedTable(track: Tracking, table: Table, depth = 0): Table {
  const out: Table = { ...table, rows: table.rows.map((row) => insertedRow(track, row, depth)) };
  delete out.srcStart;
  delete out.srcEnd;
  return out;
}

/** A cell this command puts in (with a column): it says so, and so does everything in it. */
export function insertedCell(track: Tracking, cell: TableCell, depth = 0): TableCell {
  const props: CellProps = { ...cell.props, inserted: track.revision() };
  delete props.deleted;
  return { ...cell, props, blocks: insertedBlocks(track, cell.blocks, depth) };
}

/** An empty paragraph whose mark this command inserted: what stands where a table needs a paragraph beside it. */
export function insertedFiller(track: Tracking, wrap: Wrapper | undefined): Paragraph {
  const out: Paragraph = { type: 'paragraph', props: withInsertedMark(track, EMPTY_PARA_PROPS), children: [] };
  if (wrap) out.wrap = wrap;
  return out;
}

// ------------------------------------------------------------------ ownership

/** Is everything in the blocks something this author put in since the document was opened. */
function ownBlocks(track: Tracking, blocks: Block[], depth: number): boolean {
  if (depth > MAX_DEPTH) return false;
  for (const block of blocks) {
    if (frozen(block.wrap)) return false;
    if (block.type === 'paragraph') {
      if (!isOwnParagraph(block, track)) return false;
    } else if (block.type === 'table') {
      for (const row of block.rows) if (!isOwnRow(track, row, depth + 1)) return false;
    } else return false;
  }
  return true;
}

const ownCells = (track: Tracking, cells: TableCell[], depth: number): boolean => cells.every((cell) => !frozen(cell.wrap) && ownBlocks(track, cell.blocks, depth));

/**
 * Is a row all the author's, made in this session: the row is a tracked
 * insertion of theirs and so is everything in it. Such a row is taken out
 * for good; any other is deleted as a tracked change.
 */
export function isOwnRow(track: Tracking, row: TableRow, depth = 0): boolean {
  if (!track.made(row.props.inserted) || row.props.deleted || frozen(row.wrap)) return false;
  return ownCells(track, row.cells, depth);
}

/** Is a cell all the author's, made in this session with its column. */
export function isOwnCell(track: Tracking, cell: TableCell): boolean {
  if (!track.made(cell.props.inserted) || cell.props.deleted || frozen(cell.wrap)) return false;
  return ownBlocks(track, cell.blocks, 0);
}
