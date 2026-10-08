// Text commands: typing, deleting, breaking and joining paragraphs. Each
// returns the operations that make the change and the selection it leaves;
// nothing is changed until the caller applies them (see DocHistory).

import { paragraphText, EMPTY_PARA_PROPS, type Block, type CellProps, type DocModel, type ParaProps, type Paragraph, type RunItem, type RunProps, type Table, type TableCell, type TableRow, type Wrapper } from '../../model/doc/types';
import { buildEdit, tableContextOf, type DocEdit, type DocEditOptions, type Draft } from './draft';
import { TRACK_UNSUPPORTED, afterCut, insertionLike, insertionWrap, isDeletedMark, isOwnMark, joinsWithoutLoss, strikePieces, withFreshChanges, withMarkFlagsOf, withMarkRevision, withPlainMark, type Tracking } from './track';
import { inList, patchRunProps, withoutNumbering, type RunFormatPatch } from './format';
import {
  atomFollows,
  boundary,
  caretIndex,
  cleanText,
  flatten,
  groupFor,
  harvestBlocks,
  harvestTable,
  inheritedProps,
  isAtom,
  isStaleHint,
  itemPiece,
  offsetOf,
  paragraphWith,
  rebuild,
  removePieces,
  startIndex,
  textPieces,
  wrapBetween,
  type Harvest,
  type Piece,
} from './pieces';
import {
  asRange,
  comparePositions,
  containerBlocks,
  hasAtomAt,
  inTextBox,
  nextCharOffset,
  nextWordOffset,
  orderRange,
  paragraphIn,
  paragraphLength,
  previousCharOffset,
  previousWordOffset,
  resolvePosition,
  samePosition,
  type DocPosition,
  type DocRange,
} from './position';

const bad = (message: string): never => {
  throw new RangeError(message);
};

/** A position as the caret a command leaves: before what has no width at its offset when the position the command was given was (`before`) and something of the kind is there now. */
function sided(doc: DocModel, position: DocPosition, before: boolean): DocPosition {
  if (!before) return position;
  return hasAtomAt(resolvePosition(doc, position).paragraph, position.offset) ? { ...position, before: true } : position;
}

/** Paragraph properties without the section break: of the two halves of a paragraph only the second ends the section. */
export function withoutSection(props: ParaProps): ParaProps {
  if (props.section === undefined) return props;
  const out = { ...props };
  delete out.section;
  return Object.keys(out).length ? out : EMPTY_PARA_PROPS;
}

/**
 * Properties of the paragraph two paragraphs become when the mark between
 * them goes. UNSURE: in the file format the surviving mark is the second
 * one, and old versions of Word gave the joined paragraph its formatting;
 * current Word keeps the formatting of the first paragraph unless that one is
 * empty. That is the rule here. The section break of the second paragraph
 * always stays: it is the end of the joined paragraph.
 */
export function joinedProps(first: Paragraph, last: Paragraph): ParaProps {
  if (paragraphLength(first) === 0) return last.props;
  const section = last.props.section;
  const base = withoutSection(first.props);
  return section === undefined ? base : { ...base, section };
}

/**
 * What a cleared cell holds: one paragraph, formatted like the first one the
 * cell had, with everything of the cell that is not content in it (marks of
 * bookmarks and comments, comment references, kept XML, tracked deletions,
 * from the tables inside the cell too), after the blocks of the cell that
 * are not content (marks and comments between its paragraphs). Undefined
 * when the cell is like that already.
 */
function clearedBlocks(cell: TableCell): Block[] | undefined {
  let first: Paragraph | undefined;
  let paragraphs = 0;
  let tables = 0;
  for (const block of cell.blocks) {
    if (block.type === 'paragraph') {
      if (!first) first = block;
      paragraphs++;
    } else if (block.type === 'table') tables++;
  }
  const harvest: Harvest = { pieces: [], blocks: [] };
  harvestBlocks(cell.blocks, harvest);
  const left = removePieces(harvest.pieces, 0, harvest.pieces.length);
  if (paragraphs === 1 && tables === 0 && left.length === harvest.pieces.length) return undefined;
  const paragraph: Paragraph = { type: 'paragraph', props: first ? withoutSection(first.props) : EMPTY_PARA_PROPS, children: rebuild(left) };
  harvest.blocks.push(paragraph);
  return harvest.blocks;
}

/** Grid columns a cell covers, as [start, end). */
function cellSpan(row: TableRow, index: number): [number, number] {
  let at = row.props.gridBefore || 0;
  for (let c = 0; c < index; c++) at += spanOf(row.cells[c].props);
  return [at, at + spanOf(row.cells[index].props)];
}

const spanOf = (props: CellProps): number => (props.gridSpan !== undefined && props.gridSpan > 1 ? props.gridSpan : 1);

/** A copy of a table with the content of the cells `clear` picks removed; the table itself when nothing changes. */
function tableCleared(table: Table, clear: (row: number, cell: number) => boolean): Table {
  let rows: TableRow[] | undefined;
  for (let r = 0; r < table.rows.length; r++) {
    const row = table.rows[r];
    let cells: TableCell[] | undefined;
    for (let c = 0; c < row.cells.length; c++) {
      if (!clear(r, c)) continue;
      const blocks = clearedBlocks(row.cells[c]);
      if (!blocks) continue;
      if (!cells) cells = row.cells.slice();
      cells[c] = { ...row.cells[c], blocks };
    }
    if (!cells) continue;
    if (!rows) rows = table.rows.slice();
    rows[r] = { ...row, cells };
  }
  if (!rows) return table;
  const out: Table = { ...table, rows };
  delete out.srcStart;
  delete out.srcEnd;
  return out;
}

function appendBlocks(target: Block[], blocks: Block[]): void {
  for (let i = 0; i < blocks.length; i++) target.push(blocks[i]);
}

/** Pieces of several paragraphs in a row. */
function appendPieces(target: Piece[], pieces: Piece[]): void {
  for (let i = 0; i < pieces.length; i++) target.push(pieces[i]);
}

const samePath = (a: number[] | undefined, b: number[]): boolean => a !== undefined && a.length === b.length && a.every((value, i) => value === b[i]);

/** Which cells of a table a range between two of its cells takes: all of them from the first to the last in one row, a rectangle of the grid over several rows. */
export function cellsBetween(table: Table, r1: number, c1: number, r2: number, c2: number): (row: number, cell: number) => boolean {
  if (r1 === r2) return (r, c) => r === r1 && c >= c1 && c <= c2;
  const first = cellSpan(table.rows[r1], c1);
  const last = cellSpan(table.rows[r2], c2);
  const left = Math.min(first[0], last[0]);
  const right = Math.max(first[1], last[1]);
  return (r, c) => {
    if (r < r1 || r > r2) return false;
    const span = cellSpan(table.rows[r], c);
    return span[0] < right && span[1] > left;
  };
}

/** How far the paths of two positions agree, and the level of blocks at which they part. */
export function parting(start: DocPosition, end: DocPosition): { depth: number; level: number } {
  let depth = 0;
  while (depth < start.path.length && depth < end.path.length && start.path[depth] === end.path[depth]) depth++;
  return { depth, level: depth - (depth % 3) };
}

/** Start of the first paragraph of a cell; `fallback` when the cell holds none. */
function cellStart(doc: DocModel, story: string, cell: number[], fallback: DocPosition): DocPosition {
  const path = paragraphIn(containerBlocks(doc, story, cell), cell, 0, 1);
  if (!path) return fallback;
  if (samePath(path, fallback.path) && fallback.offset === 0) return fallback;
  return hasAtomAt(resolvePosition(doc, { story, path, offset: 0 }).paragraph, 0) ? { story, path, offset: 0, before: true } : { story, path, offset: 0 };
}

/** End of the last paragraph of a cell. */
function cellEnd(doc: DocModel, story: string, cell: number[], fallback: DocPosition): DocPosition {
  const blocks = containerBlocks(doc, story, cell);
  const path = paragraphIn(blocks, cell, blocks.length - 1, -1);
  if (!path) return fallback;
  const offset = resolvePosition(doc, { story, path, offset: 0 }).length;
  if (samePath(path, fallback.path) && fallback.offset === offset && fallback.before !== true) return fallback;
  return { story, path, offset };
}

/**
 * The range a deletion really covers. A range that crosses the border of a
 * table takes whole cells, as a selection in Word does: where it begins in a
 * table and ends outside of it (at that level), it covers the cell it begins
 * in from its start; where it ends in a table it began outside of, the cell
 * it ends in to its end; between two cells of one table, both cells whole.
 * deleteRange and insertText over a range delete exactly this, so an editor
 * shows this range as the selection. Over several rows of one table the
 * cells that are emptied are those of the rectangle between the two cells
 * (by grid column), which a range cannot say: the range returned runs from
 * the first emptied cell of the first row to the last one of the last row.
 *
 * Returns the range itself when nothing changes. Throws RangeError for a
 * range that is not one of the document.
 */
export function expandRangeForTables(doc: DocModel, range: DocRange): DocRange {
  const at = asRange(range);
  const { start, end } = orderRange(at);
  resolvePosition(doc, start);
  resolvePosition(doc, end);
  if (inTextBox(start.path) || inTextBox(end.path)) return range;
  const { depth, level } = parting(start, end);
  if (depth === start.path.length && depth === end.path.length) return range;
  const story = start.story;
  let from = start;
  let to = end;
  if (depth % 3 !== 0) {
    const table = containerBlocks(doc, story, start.path.slice(0, level))[start.path[level]] as Table;
    const [r1, c1, r2, c2] = [start.path[level + 1], start.path[level + 2], end.path[level + 1], end.path[level + 2]];
    const takes = cellsBetween(table, r1, c1, r2, c2);
    let first = c1;
    let last = c2;
    if (r1 !== r2) {
      const top = table.rows[r1].cells.length;
      first = 0;
      while (first < top - 1 && !takes(r1, first)) first++;
      last = table.rows[r2].cells.length - 1;
      while (last > 0 && !takes(r2, last)) last--;
    }
    const base = start.path.slice(0, level + 1);
    from = cellStart(doc, story, base.concat([r1, first]), start);
    to = cellEnd(doc, story, base.concat([r2, last]), end);
  } else {
    if (start.path.length > level + 1) from = cellStart(doc, story, start.path.slice(0, level + 3), start);
    if (end.path.length > level + 1) to = cellEnd(doc, story, end.path.slice(0, level + 3), end);
  }
  if (from === start && to === end) return range;
  return start === at.anchor ? { anchor: from, focus: to } : { anchor: to, focus: from };
}

/** What kind of content block comes next before (step -1) or after (1) a block of a list; blocks without content are passed over. */
export function neighbour(blocks: Block[], index: number, step: 1 | -1): number {
  for (let i = index + step, tries = 0; i >= 0 && i < blocks.length && tries < 64; i += step, tries++) {
    if (blocks[i].type === 'paragraph' || blocks[i].type === 'table') return i;
  }
  return -1;
}

/**
 * Removes what lies between two positions (`start` before `end`) and returns
 * where the caret is afterwards.
 *
 * Within a paragraph the content goes. Across paragraphs the first one keeps
 * what is before the range and takes what the last one has after it; the
 * paragraphs and tables in between go. Where the range ends inside a table
 * it began outside of (or the other way round), and between cells of one
 * table, the cells in the range are emptied, whole (see
 * expandRangeForTables), and the table stays as it is. A table goes when
 * the range covers all of it: when it lies between the ends of the range, or
 * the range begins at its very start and ends after it, or begins before it
 * and ends at its very end (so "select all, delete" leaves no empty tables).
 * UNSURE: Word deletes the rows a selection that leaves the table touches;
 * here they are emptied unless the whole table is in the range.
 *
 * What is not content (bookmarks, comment marks and references, kept XML,
 * tracked deletions, fields the range only partly covers) is never removed:
 * of removed paragraphs and tables it moves into the paragraph that stays,
 * of an emptied cell into the paragraph the cell keeps. The list of blocks
 * always keeps a paragraph where one is needed: after a table at its end,
 * between two tables, in a list that would be empty.
 */
export function deleteSpan(draft: Draft, start: DocPosition, end: DocPosition): DocPosition {
  const doc = draft.doc;
  const story = start.story;
  const a = resolvePosition(doc, start);
  const b = resolvePosition(doc, end);
  if (inTextBox(start.path) || inTextBox(end.path)) bad('The content of a text box cannot be edited');
  if (comparePositions(start, end) > 0) bad('The range is the wrong way round');
  const sideOfStart = start.before === true;
  if (draft.track) return sided(doc, strikeSpan(draft, draft.track, start, end), sideOfStart);

  const { depth, level } = parting(start, end);
  if (depth === start.path.length && depth === end.path.length) {
    const pieces = flatten(a.paragraph);
    const from = startIndex(pieces, start);
    const to = caretIndex(pieces, end);
    const offset = offsetOf(pieces, from);
    if (from < to) {
      const left = removePieces(pieces, from, to);
      if (left.length !== pieces.length) draft.setParagraph(story, start.path, paragraphWith(a.paragraph, rebuild(left)));
    }
    return sided(doc, { story, path: start.path.slice(), offset }, sideOfStart);
  }

  const container = start.path.slice(0, level);
  if (depth % 3 !== 0) {
    // Two cells of one table: the cells between them are emptied.
    const blocks = containerBlocks(doc, story, container);
    const table = blocks[start.path[level]] as Table;
    const cleared = tableCleared(table, cellsBetween(table, start.path[level + 1], start.path[level + 2], end.path[level + 1], end.path[level + 2]));
    if (cleared !== table) draft.replace(story, container, start.path[level], 1, [cleared]);
    const cell = containerBlocks(doc, story, start.path.slice(0, level + 3));
    const path = paragraphIn(cell, start.path.slice(0, level + 3), 0, 1);
    return path ? { story, path, offset: 0 } : { story, path: start.path.slice(), offset: 0 };
  }

  const blocks = containerBlocks(doc, story, container);
  const first = start.path[level];
  const last = end.path[level];
  const startPara = start.path.length === level + 1 ? a.paragraph : undefined;
  const endPara = end.path.length === level + 1 ? b.paragraph : undefined;
  const startTable = startPara ? undefined : (blocks[first] as Table);
  const endTable = endPara ? undefined : (blocks[last] as Table);
  // A table the range covers from its very first position, or to its very last, goes as a whole.
  const startGoes = startTable !== undefined && start.offset === 0 && samePath(paragraphIn(blocks, container, first, 1), start.path);
  const endGoes = endTable !== undefined && end.offset === b.length && samePath(paragraphIn(blocks, container, last, -1), end.path);

  // What stays of the paragraphs and tables in the range, in order, and the blocks without content between them.
  const harvest: Harvest = { pieces: [], blocks: [] };
  const pieces = harvest.pieces;
  const kept = harvest.blocks;
  let from = 0;
  let caretOffset = 0;
  if (startPara) {
    appendPieces(pieces, flatten(startPara));
    from = startIndex(pieces, start);
    caretOffset = offsetOf(pieces, from);
  } else if (startGoes) harvestTable(startTable as Table, harvest);
  let middle: Paragraph | undefined;
  let tableBetween = false;
  for (let i = first + 1; i < last; i++) {
    const block = blocks[i];
    if (block.type === 'paragraph') {
      appendPieces(pieces, flatten(block));
      middle = block;
    } else if (block.type === 'table') {
      harvestTable(block, harvest);
      tableBetween = true;
    } else kept.push(block);
  }
  if (endGoes) harvestTable(endTable as Table, harvest);
  let to = pieces.length;
  if (endPara) {
    const tail = flatten(endPara);
    to += caretIndex(tail, end);
    appendPieces(pieces, tail);
  }
  const children = rebuild(removePieces(pieces, from, to));

  const out: Block[] = [];
  /** Index in `out` of the paragraph that takes what stays of the range. */
  let holder = -1;
  const startStays = startTable !== undefined && !startGoes;
  const endStays = endTable !== undefined && !endGoes;
  if (startStays) {
    const [r1, c1] = [start.path[level + 1], start.path[level + 2]];
    out.push(tableCleared(startTable as Table, (r, c) => r > r1 || (r === r1 && c >= c1)));
  }
  if (startPara) {
    holder = out.length;
    out.push(paragraphWith(startPara, children, endPara ? joinedProps(startPara, endPara) : startPara.props));
    appendBlocks(out, kept);
  } else if (endPara) {
    appendBlocks(out, kept);
    holder = out.length;
    out.push(paragraphWith(endPara, children, endPara.props));
  } else {
    // From a table to a table. A paragraph that was between them stays, empty; without one, a paragraph is
    // made where the list would otherwise end with a table, have two tables touch, or hold nothing to stand in.
    const before = neighbour(blocks, first, -1);
    const after = neighbour(blocks, last, 1);
    const tableBefore = before >= 0 && blocks[before].type === 'table';
    const tableAfter = after >= 0 && blocks[after].type === 'table';
    let needed: boolean;
    if (middle || children.length > 0) needed = true;
    else if (startStays && endStays) needed = tableBetween;
    else if (startStays) needed = after < 0 || tableAfter;
    else if (endStays) needed = tableBefore;
    else needed = after < 0 ? before < 0 || tableBefore : tableBefore && tableAfter;
    if (needed) {
      holder = out.length;
      out.push(middle ? paragraphWith(middle, children, withoutSection(middle.props), false) : { type: 'paragraph', props: EMPTY_PARA_PROPS, children });
    }
    appendBlocks(out, kept);
  }
  if (endStays) {
    const [r2, c2] = [end.path[level + 1], end.path[level + 2]];
    out.push(tableCleared(endTable as Table, (r, c) => r < r2 || (r === r2 && c <= c2)));
  }
  draft.replace(story, container, first, last - first + 1, out);

  if (startPara) return sided(doc, { story, path: start.path.slice(), offset: caretOffset }, sideOfStart);
  if (startStays) {
    const cell = start.path.slice(0, level + 3);
    const path = paragraphIn(containerBlocks(doc, story, cell), cell, 0, 1);
    return path ? { story, path, offset: 0 } : { story, path: start.path.slice(), offset: 0 };
  }
  // The table the range began in is gone: the caret goes to what took its place, or to what follows, or to what is before.
  if (holder >= 0) return { story, path: container.concat([first + holder]), offset: 0 };
  const now = containerBlocks(doc, story, container);
  const next = paragraphIn(now, container, first, 1);
  if (next) return { story, path: next, offset: 0 };
  const previous = paragraphIn(now, container, first - 1, -1);
  if (!previous) return bad('The story holds no paragraph');
  return { story, path: previous, offset: resolvePosition(doc, { story, path: previous, offset: 0 }).length };
}

/** The properties of the first paragraph after block `index`, up to block `last` (which is one). */
function nextParagraphProps(blocks: Block[], index: number, last: number): ParaProps {
  for (let i = index + 1; i <= last; i++) {
    const block = blocks[i];
    if (block.type === 'paragraph') return block.props;
  }
  return EMPTY_PARA_PROPS;
}

/**
 * deleteSpan() where changes are tracked. What the range holds stays, as
 * tracked deletions, and so do its paragraphs: the mark of each paragraph
 * the range runs past says that it was deleted. What the author of the
 * change had put directly into a tracked insertion goes for good, text and,
 * between paragraphs formatted alike, paragraph marks. A range that takes cells of a table, or a table, is refused:
 * deleting those is a change of the table that is not tracked yet.
 */
function strikeSpan(draft: Draft, track: Tracking, start: DocPosition, end: DocPosition): DocPosition {
  const doc = draft.doc;
  const story = start.story;
  const a = resolvePosition(doc, start);
  const { depth, level } = parting(start, end);
  if (depth === start.path.length && depth === end.path.length) {
    const pieces = flatten(a.paragraph);
    const from = startIndex(pieces, start);
    const to = caretIndex(pieces, end);
    const offset = offsetOf(pieces, from);
    const struck = strikePieces(track, pieces, from, to);
    if (struck.changed) draft.setParagraph(story, start.path, paragraphWith(a.paragraph, rebuild(struck.pieces)));
    return { story, path: start.path.slice(), offset };
  }
  if (depth % 3 !== 0 || start.path.length !== level + 1 || end.path.length !== level + 1) bad(TRACK_UNSUPPORTED);
  const container = start.path.slice(0, level);
  const blocks = containerBlocks(doc, story, container);
  const first = start.path[level];
  const last = end.path[level];
  for (let i = first + 1; i < last; i++) if (blocks[i].type === 'table') bad(TRACK_UNSUPPORTED);

  const out: Block[] = [];
  /** Blocks without content that follow a paragraph whose mark went for good: they follow the paragraph it joined. */
  let loose: Block[] = [];
  /** The paragraphs joined so far to the one being read: the first of them, and their pieces. */
  let open: { first: Paragraph; pieces: Piece[] } | undefined;
  let caretOffset = 0;
  let any = false;
  for (let i = first; i <= last; i++) {
    const block = blocks[i];
    if (block.type !== 'paragraph') {
      (open ? loose : out).push(block);
      continue;
    }
    const pieces = flatten(block);
    let from = 0;
    let to = pieces.length;
    if (i === first) {
      from = startIndex(pieces, start);
      caretOffset = offsetOf(pieces, from);
      to = pieces.length;
    }
    if (i === last) to = caretIndex(pieces, end);
    const struck = strikePieces(track, pieces, from, to);
    let changed = struck.changed;
    let content = struck.pieces;
    if (open) {
      content = open.pieces.concat(afterCut(open.pieces, content, track));
      changed = true;
    }
    if (i < last && isOwnMark(block.props, track) && joinsWithoutLoss(block.props, nextParagraphProps(blocks, i, last))) {
      // The mark was a tracked insertion of this author, between two paragraphs formatted alike: it goes, and the
      // next paragraph continues this one. Between paragraphs that differ it is deleted like any other mark, so
      // that the second keeps its formatting.
      open = { first: open ? open.first : block, pieces: content };
      any = true;
      continue;
    }
    let props = block.props;
    // The mark that stays is the mark of this paragraph, whatever the paragraphs joined to it said of theirs.
    if (open) props = withMarkFlagsOf(joinedProps(open.first, block), block.props);
    if (i < last && !isDeletedMark(props)) {
      props = withMarkRevision(props, 'deleted', track.revision());
      changed = true;
    }
    if (changed) {
      any = true;
      out.push(paragraphWith(open ? open.first : block, rebuild(content), props));
    } else out.push(block);
    appendBlocks(out, loose);
    loose = [];
    open = undefined;
  }
  if (any) draft.replace(story, container, first, last - first + 1, out);
  return { story, path: start.path.slice(), offset: caretOffset };
}

/**
 * Deletes the content of a range. The selection left is a caret where the
 * range began. A range that crosses the border of a table deletes what
 * expandRangeForTables() returns for it.
 */
export function deleteRange(doc: DocModel, range: DocRange, options?: DocEditOptions): DocEdit {
  const at = asRange(range);
  const build = (draft: Draft) => {
    const { start, end } = orderRange(at);
    resolvePosition(doc, start);
    resolvePosition(doc, end);
    return samePosition(start, end) ? start : deleteSpan(draft, start, end);
  };
  return buildEdit(doc, build, options, true);
}

/** Most paragraphs one insertion makes: a paste beyond this is refused rather than left half done. */
const MAX_INSERTED_PARAGRAPHS = 1000000;

function insertAt(draft: Draft, position: DocPosition, text: string, format: RunFormatPatch | undefined): DocPosition {
  const doc = draft.doc;
  const at = resolvePosition(doc, position);
  if (inTextBox(position.path)) bad('The content of a text box cannot be edited');
  const paragraph = at.paragraph;
  const pieces = flatten(paragraph);
  const index = caretIndex(pieces, position);
  const offset = offsetOf(pieces, index);
  const inherited = inheritedProps(pieces, index, paragraph.props.markProps);
  const props = format ? patchRunProps(inherited, format, doc) : inherited;
  const wrap = wrapBetween(index > 0 ? pieces[index - 1].wrap : undefined, index < pieces.length ? pieces[index].wrap : undefined);
  const lines = text.split('\n');
  if (lines.length > MAX_INSERTED_PARAGRAPHS) bad('The text has too many paragraphs');
  const story = position.story;
  const side = position.before === true;
  if (draft.track) return sided(doc, insertTracked(draft, draft.track, position, paragraph, pieces, index, lines, props, wrap), side);
  const group = groupFor(pieces, index, props, wrap);
  const head = pieces.slice(0, index);
  appendPieces(head, textPieces(lines[0], props, wrap, group));
  if (lines.length === 1) {
    for (let i = index; i < pieces.length; i++) head.push(pieces[i]);
    draft.setParagraph(story, position.path, paragraphWith(paragraph, rebuild(head)));
    return sided(doc, { story, path: position.path.slice(), offset: offset + lines[0].length }, side);
  }
  // Every line but the last ends a paragraph formatted like this one; the section break stays with the last.
  const inner = withoutSection(paragraph.props);
  const out: Block[] = [paragraphWith(paragraph, rebuild(head), inner)];
  for (let i = 1; i < lines.length - 1; i++) out.push(paragraphWith(paragraph, rebuild(textPieces(lines[i], props, wrap, {})), inner, false));
  const lastLine = lines[lines.length - 1];
  const tail = textPieces(lastLine, props, wrap, index < pieces.length ? groupFor(pieces, index, props, wrap) : {});
  for (let i = index; i < pieces.length; i++) tail.push(pieces[i]);
  out.push(paragraphWith(paragraph, rebuild(tail), paragraph.props, false));
  const container = position.path.slice(0, -1);
  const blockIndex = position.path[position.path.length - 1];
  draft.replace(story, container, blockIndex, 1, out);
  return sided(doc, { story, path: container.concat([blockIndex + lines.length - 1]), offset: lastLine.length }, side);
}

/**
 * insertAt() where changes are tracked: the text goes into a tracked
 * insertion (the one of the same author beside it, when there is one), and
 * the mark of every paragraph the text ends says that it was inserted. In
 * the file the mark that is new is the one of the first half of a paragraph
 * that was cut, so that is where it is said.
 */
function insertTracked(draft: Draft, track: Tracking, position: DocPosition, paragraph: Paragraph, source: Piece[], index: number, lines: string[], props: RunProps, base: Wrapper | undefined): DocPosition {
  const story = position.story;
  const offset = offsetOf(source, index);
  const placed = insertionWrap(track, source, index, base);
  const pieces = placed.pieces;
  const wrap = placed.wrap;
  const head = pieces.slice(0, index);
  appendPieces(head, textPieces(lines[0], props, wrap, groupFor(pieces, index, props, wrap)));
  if (lines.length === 1) {
    // A deletion or a move the text went into the middle of is two elements now: the second gets an id of its own.
    appendPieces(head, afterCut(pieces.slice(0, index), pieces.slice(index), track, wrap));
    draft.setParagraph(story, position.path, paragraphWith(paragraph, rebuild(head)));
    return { story, path: position.path.slice(), offset: offset + lines[0].length };
  }
  const inner = withPlainMark(withoutSection(paragraph.props));
  const ended = (): ParaProps => withMarkRevision(withFreshChanges(inner, track), 'inserted', track.revision());
  const out: Block[] = [paragraphWith(paragraph, rebuild(head), ended())];
  for (let i = 1; i < lines.length - 1; i++) {
    const made = lines[i] === '' ? [] : textPieces(lines[i], props, insertionLike(track, wrap), {});
    out.push(paragraphWith(paragraph, rebuild(made), ended(), false));
  }
  const lastLine = lines[lines.length - 1];
  const tail = lastLine === '' ? [] : textPieces(lastLine, props, insertionLike(track, wrap), {});
  // What follows the caret is in other elements now than what is before it.
  appendPieces(tail, afterCut(head, pieces.slice(index), track));
  out.push(paragraphWith(paragraph, rebuild(tail), paragraph.props, false));
  const container = position.path.slice(0, -1);
  const blockIndex = position.path[position.path.length - 1];
  draft.replace(story, container, blockIndex, 1, out);
  return { story, path: container.concat([blockIndex + lines.length - 1]), offset: lastLine.length };
}

/**
 * Types or pastes text at a position, or over a range (whose content goes
 * first). "\n" ends a paragraph, a tab character becomes a tab and a vertical
 * tab ("\v", Shift+Enter) a line break. The text takes the formatting of the
 * text to its left, at the start of a paragraph of the text to its right, in
 * an empty paragraph of the paragraph mark (see inheritedProps for what text
 * beside a hyperlink takes); `format` is laid over that (the "bold is on" of
 * a caret that has no text yet).
 */
export function insertText(doc: DocModel, at: DocPosition | DocRange, text: string, format?: RunFormatPatch, options?: DocEditOptions): DocEdit {
  if (typeof text !== 'string') bad('The text to insert must be a string');
  const range = asRange(at);
  const clean = cleanText(text);
  const pending = format === null ? undefined : format;
  const build = (draft: Draft) => {
    const { start, end } = orderRange(range);
    resolvePosition(doc, start);
    resolvePosition(doc, end);
    if (pending !== undefined) patchRunProps({}, pending, doc);
    const typing = samePosition(start, end);
    const caret = typing ? start : deleteSpan(draft, start, end);
    if (clean === '') return caret;
    const after = insertAt(draft, caret, clean, pending);
    const selection = { anchor: after, focus: after };
    // One character typed at a caret: joins the characters typed before it into one undo step.
    if (!typing || nextCharOffset(clean, 0) !== clean.length || clean === '\n' || clean === '\t' || clean === '\v') return { selection };
    const space = clean === ' ' || clean === ' ';
    return { selection, mergeKey: 'type:' + start.story + ':' + start.path.join(','), seal: space };
  };
  return buildEdit(doc, build, options, true);
}

/** Joins the paragraph at `second` of a list of blocks onto the one at `first`; what lies between them (marks, kept XML) follows the joined paragraph. */
function joinParagraphs(draft: Draft, story: string, container: number[], first: number, second: number): DocPosition {
  const blocks = containerBlocks(draft.doc, story, container);
  const a = blocks[first] as Paragraph;
  const b = blocks[second] as Paragraph;
  const pieces = flatten(a);
  appendPieces(pieces, flatten(b));
  const out: Block[] = [paragraphWith(a, rebuild(pieces), joinedProps(a, b))];
  for (let i = first + 1; i < second; i++) out.push(blocks[i]);
  draft.replace(story, container, first, second - first + 1, out);
  return { story, path: container.concat([first]), offset: paragraphLength(a) };
}

/**
 * joinParagraphs() where changes are tracked. The two paragraphs stay two:
 * the mark of the first says that it was deleted. A mark this author had
 * inserted as a tracked change goes for good instead, and the paragraphs are
 * joined. The caret is left where the deletion went: at the end of the
 * first paragraph for Backspace, at the start of the second for Delete, also
 * when the mark was deleted already.
 */
function strikeMark(draft: Draft, track: Tracking, story: string, container: number[], first: number, second: number, forward: boolean): DocPosition {
  const blocks = containerBlocks(draft.doc, story, container);
  const a = blocks[first] as Paragraph;
  const b = blocks[second] as Paragraph;
  const length = paragraphLength(a);
  if (isOwnMark(a.props, track) && joinsWithoutLoss(a.props, b.props)) {
    const head = flatten(a);
    const pieces = head.concat(afterCut(head, flatten(b), track));
    const out: Block[] = [paragraphWith(a, rebuild(pieces), withMarkFlagsOf(joinedProps(a, b), b.props))];
    for (let i = first + 1; i < second; i++) out.push(blocks[i]);
    draft.replace(story, container, first, second - first + 1, out);
    return { story, path: container.concat([first]), offset: length };
  }
  if (!isDeletedMark(a.props)) draft.setParagraph(story, container.concat([first]), paragraphWith(a, a.children, withMarkRevision(a.props, 'deleted', track.revision())));
  return forward ? { story, path: container.concat([second]), offset: 0 } : { story, path: container.concat([first]), offset: length };
}

/**
 * Removes the pieces [from, to) of a paragraph and returns the caret that
 * leaves. The pieces of `hints` go too when anything is removed: the stale
 * hints Backspace or Delete passed on its way to what it takes.
 */
function removeFrom(draft: Draft, position: DocPosition, paragraph: Paragraph, pieces: Piece[], from: number, to: number, hints: Set<Piece>): DocPosition {
  const offset = offsetOf(pieces, from);
  if (draft.track) {
    const struck = strikePieces(draft.track, pieces, from, to);
    if (struck.changed) {
      const kept = hints.size ? struck.pieces.filter((piece) => !hints.has(piece)) : struck.pieces;
      draft.setParagraph(position.story, position.path, paragraphWith(paragraph, rebuild(kept)));
    }
    return { story: position.story, path: position.path.slice(), offset };
  }
  let left = removePieces(pieces, from, to);
  if (left.length !== pieces.length) {
    if (hints.size) left = left.filter((piece) => !hints.has(piece));
    draft.setParagraph(position.story, position.path, paragraphWith(paragraph, rebuild(left)));
  }
  return { story: position.story, path: position.path.slice(), offset };
}

/** The stale hints among the pieces [from, to): pieces without width are never cut, so they stay the objects they are while the list changes. */
function hintsBetween(pieces: Piece[], from: number, to: number): Set<Piece> {
  const out = new Set<Piece>();
  for (let q = from; q < to; q++) if (isStaleHint(pieces[q])) out.add(pieces[q]);
  return out;
}

/** The pieces a lock holds around the piece at `index`, as [from, to). */
function lockSpan(pieces: Piece[], index: number): [number, number] {
  const lock = pieces[index].lock;
  let from = index;
  let to = index + 1;
  while (from > 0 && pieces[from - 1].lock === lock) from--;
  while (to < pieces.length && pieces[to].lock === lock) to++;
  return [from, to];
}

const checkUnit = (unit: unknown): void => {
  if (unit !== 'char' && unit !== 'word') bad('The unit of a deletion is "char" or "word"');
};

/**
 * Backspace. Removes the character (or the word) before the caret; a drawing,
 * a note reference or a computed field there goes as one unit. What does not
 * show is never the unit that goes: the caret passes over it to what shows.
 * Of what it passes, marks of bookmarks and comments and a soft hyphen stay,
 * and the hints of the last layout and spelling check (w:lastRenderedPageBreak,
 * w:proofErr) are dropped, being stale once the text changes. UNSURE: a soft
 * hyphen takes no offset (paragraphText leaves it out), so no caret can stand
 * beside it and only a range over it removes it.
 *
 * At the start of a paragraph, in Word's order: an item of a list leaves the
 * list; else an indent goes (first the first line indent, then the left one);
 * else the paragraph joins the one before it. Nothing happens at the start of
 * a table cell or of the story, and after a table.
 */
export function deleteBackward(doc: DocModel, position: DocPosition, unit: 'char' | 'word' = 'char', options?: DocEditOptions): DocEdit {
  checkUnit(unit);
  const build = (draft: Draft) => {
    const at = resolvePosition(doc, position);
    if (inTextBox(position.path)) bad('The content of a text box cannot be edited');
    const paragraph = at.paragraph;
    const side = position.before === true;
    const pieces = flatten(paragraph);
    const index = caretIndex(pieces, position);
    let k = index - 1;
    while (k >= 0 && pieces[k].len === 0 && !isAtom(pieces[k]) && pieces[k].lock === undefined) k--;
    if (k >= 0) {
      const hints = hintsBetween(pieces, k + 1, index);
      if (pieces[k].lock !== undefined) {
        const [from, to] = lockSpan(pieces, k);
        return sided(doc, removeFrom(draft, position, paragraph, pieces, from, to, hints), side);
      }
      if (isAtom(pieces[k])) return sided(doc, removeFrom(draft, position, paragraph, pieces, k, k + 1, hints), side);
      const text = paragraphText(paragraph);
      const end = offsetOf(pieces, k + 1);
      const begin = unit === 'word' ? previousWordOffset(text, end) : previousCharOffset(text, end);
      const from = boundary(pieces, begin, 'before');
      const to = boundary(pieces, end, 'after');
      return sided(doc, removeFrom(draft, position, paragraph, pieces, from, to, hints), side);
    }
    const track = draft.track;
    if (track) {
      // Leaving a list and taking back an indent are changes of formatting, which are not tracked: the mark before is deleted.
      const previous = neighbour(at.blocks, at.index, -1);
      if (previous < 0 || at.blocks[previous].type !== 'paragraph') return position;
      return sided(doc, strikeMark(draft, track, position.story, at.container, previous, at.index, false), side);
    }
    const where = tableContextOf(doc, position.story, position.path);
    if (inList(doc, paragraph.props, where)) {
      draft.setParagraph(position.story, position.path, paragraphWith(paragraph, paragraph.children, withoutNumbering(doc, paragraph.props, where)));
      return position;
    }
    const indent = paragraph.props.indent;
    if (indent && ((indent.firstLine || 0) > 0 || (indent.left || 0) > 0 || (indent.hanging || 0) > 0)) {
      const next = { ...indent };
      if ((indent.firstLine || 0) > 0) {
        delete next.firstLine;
        delete next.firstLineChars;
      } else {
        delete next.left;
        delete next.leftChars;
        delete next.hanging;
        delete next.hangingChars;
      }
      const props: ParaProps = { ...paragraph.props, indent: next };
      if (Object.keys(next).length === 0) delete props.indent;
      draft.setParagraph(position.story, position.path, paragraphWith(paragraph, paragraph.children, Object.keys(props).length ? props : EMPTY_PARA_PROPS));
      return position;
    }
    const before = neighbour(at.blocks, at.index, -1);
    if (before < 0 || at.blocks[before].type !== 'paragraph') return position;
    return sided(doc, joinParagraphs(draft, position.story, at.container, before, at.index), side);
  };
  return buildEdit(doc, build, options, true);
}

/**
 * Delete. Removes the character (or the word) after the caret, with the same
 * units as deleteBackward and the same treatment of what does not show; at
 * the end of a paragraph the next paragraph joins this one. Nothing happens
 * at the end of a table cell or of the story, and before a table. The caret
 * stays at the offset it had.
 */
export function deleteForward(doc: DocModel, position: DocPosition, unit: 'char' | 'word' = 'char', options?: DocEditOptions): DocEdit {
  checkUnit(unit);
  const build = (draft: Draft) => {
    const at = resolvePosition(doc, position);
    if (inTextBox(position.path)) bad('The content of a text box cannot be edited');
    const paragraph = at.paragraph;
    const side = position.before === true;
    const pieces = flatten(paragraph);
    const index = caretIndex(pieces, position);
    let k = index;
    while (k < pieces.length && pieces[k].len === 0 && !isAtom(pieces[k]) && pieces[k].lock === undefined) k++;
    if (k < pieces.length) {
      const hints = hintsBetween(pieces, index, k);
      if (pieces[k].lock !== undefined) {
        const [from, to] = lockSpan(pieces, k);
        return sided(doc, removeFrom(draft, position, paragraph, pieces, from, to, hints), side);
      }
      if (isAtom(pieces[k])) return sided(doc, removeFrom(draft, position, paragraph, pieces, k, k + 1, hints), side);
      const text = paragraphText(paragraph);
      const begin = offsetOf(pieces, k);
      const end = unit === 'word' ? nextWordOffset(text, begin) : nextCharOffset(text, begin);
      const from = boundary(pieces, begin, 'before');
      const to = boundary(pieces, end, 'after');
      const caret = removeFrom(draft, position, paragraph, pieces, from, to, hints);
      return sided(doc, { ...caret, offset: position.offset }, side);
    }
    const after = neighbour(at.blocks, at.index, 1);
    if (after < 0 || at.blocks[after].type !== 'paragraph') return position;
    if (draft.track) return strikeMark(draft, draft.track, position.story, at.container, at.index, after, true);
    return sided(doc, joinParagraphs(draft, position.story, at.container, at.index, after), side);
  };
  return buildEdit(doc, build, options, true);
}

/**
 * Enter. The paragraph is cut in two at the position; both halves keep its
 * style and its direct formatting, and a section break it carries stays with
 * the second. In an empty list item the paragraph leaves the list instead.
 *
 * At the end of a paragraph whose style names a style for the paragraph that
 * follows (a heading), the new paragraph takes that style and none of the
 * direct paragraph formatting (alignment, indents, spacing): in Word applying
 * the following style starts the new paragraph from the style alone. UNSURE:
 * this is Word's behaviour as documented and observed for headings; it was
 * not compared with Word for every kind of direct formatting (borders, tabs).
 */
export function splitParagraph(doc: DocModel, position: DocPosition, options?: DocEditOptions): DocEdit {
  const build = (draft: Draft) => {
    const at = resolvePosition(doc, position);
    if (inTextBox(position.path)) bad('The content of a text box cannot be edited');
    const paragraph = at.paragraph;
    const track = draft.track;
    if (track) {
      // Leaving a list and taking the style that follows a heading are changes of formatting, which are not
      // tracked: both halves keep what the paragraph had, and the mark of the first says that it was inserted.
      const all = flatten(paragraph);
      const cut = caretIndex(all, position);
      const list = position.path.slice(0, -1);
      const nth = position.path[position.path.length - 1];
      const head = all.slice(0, cut);
      draft.replace(position.story, list, nth, 1, [
        paragraphWith(paragraph, rebuild(head), withMarkRevision(withFreshChanges(withPlainMark(withoutSection(paragraph.props)), track), 'inserted', track.revision())),
        paragraphWith(paragraph, rebuild(afterCut(head, all.slice(cut), track)), paragraph.props, false),
      ]);
      return sided(doc, { story: position.story, path: list.concat([nth + 1]), offset: 0 }, position.before === true);
    }
    const where = tableContextOf(doc, position.story, position.path);
    if (at.length === 0 && inList(doc, paragraph.props, where)) {
      draft.setParagraph(position.story, position.path, paragraphWith(paragraph, paragraph.children, withoutNumbering(doc, paragraph.props, where)));
      return position;
    }
    const pieces = flatten(paragraph);
    const index = caretIndex(pieces, position);
    let second = paragraph.props;
    // At the very end: nothing that shows goes into the new paragraph.
    if (at.length > 0 && offsetOf(pieces, index) === at.length && !atomFollows(pieces, index)) {
      const styles = doc.styles;
      const own = paragraph.props.style !== undefined ? paragraph.props.style : styles.defaultIds.paragraph;
      const style = own === undefined ? undefined : styles.styles.get(own);
      const next = style && style.type === 'paragraph' && style.next !== undefined && style.next !== style.id ? styles.styles.get(style.next) : undefined;
      if (next && next.type === 'paragraph') {
        second = next.id === styles.defaultIds.paragraph ? {} : { style: next.id };
        if (paragraph.props.section) second.section = paragraph.props.section;
        if (Object.keys(second).length === 0) second = EMPTY_PARA_PROPS;
      }
    }
    const container = position.path.slice(0, -1);
    const blockIndex = position.path[position.path.length - 1];
    // The second half is a new paragraph: it does not take the identity (w14:paraId) of the first.
    draft.replace(position.story, container, blockIndex, 1, [
      paragraphWith(paragraph, rebuild(pieces.slice(0, index)), withoutSection(paragraph.props)),
      paragraphWith(paragraph, rebuild(pieces.slice(index)), second, false),
    ]);
    return sided(doc, { story: position.story, path: container.concat([blockIndex + 1]), offset: 0 }, position.before === true);
  };
  return buildEdit(doc, build, options, true);
}

const PAGE_BREAK_ITEM: RunItem = Object.freeze({ type: 'break' as const, kind: 'page' as const });

/**
 * Ctrl+Enter. A page break goes in at the position and the paragraph ends
 * after it: what follows starts the next page as a paragraph of its own.
 * Refused inside a table cell. UNSURE: Word does not put a page break into a
 * cell; it splits the table at the row instead, which is not done here.
 */
export function insertPageBreak(doc: DocModel, position: DocPosition, options?: DocEditOptions): DocEdit {
  const build = (draft: Draft) => {
    const at = resolvePosition(doc, position);
    if (inTextBox(position.path)) bad('The content of a text box cannot be edited');
    if (position.path.length > 1) bad('A page break cannot be put into a table cell');
    const paragraph = at.paragraph;
    let pieces = flatten(paragraph);
    const index = caretIndex(pieces, position);
    const props = inheritedProps(pieces, index, paragraph.props.markProps);
    let wrap = wrapBetween(index > 0 ? pieces[index - 1].wrap : undefined, index < pieces.length ? pieces[index].wrap : undefined);
    const track = draft.track;
    if (track) {
      // The break is a tracked insertion, and so is the mark of the paragraph it ends.
      const placed = insertionWrap(track, pieces, index, wrap);
      pieces = placed.pieces;
      wrap = placed.wrap;
    }
    const head = pieces.slice(0, index);
    head.push(itemPiece(PAGE_BREAK_ITEM, props, wrap, groupFor(pieces, index, props, wrap)));
    const container = position.path.slice(0, -1);
    const blockIndex = position.path[position.path.length - 1];
    const firstProps = withoutSection(paragraph.props);
    draft.replace(position.story, container, blockIndex, 1, [
      paragraphWith(paragraph, rebuild(head), track ? withMarkRevision(withFreshChanges(withPlainMark(firstProps), track), 'inserted', track.revision()) : firstProps),
      paragraphWith(paragraph, rebuild(track ? afterCut(head, pieces.slice(index), track) : pieces.slice(index)), paragraph.props, false),
    ]);
    return sided(doc, { story: position.story, path: container.concat([blockIndex + 1]), offset: 0 }, position.before === true);
  };
  return buildEdit(doc, build, options, true);
}

/**
 * Puts a block at a position: before the paragraph when the position is at
 * its start, after it at its end, and between its two halves otherwise. A
 * table never ends a list of blocks and never touches another table: an empty
 * paragraph is added where that would happen. Returns the index of the block
 * in its list.
 */
export function placeBlock(draft: Draft, position: DocPosition, block: Block): number {
  const doc = draft.doc;
  const at = resolvePosition(doc, position);
  if (inTextBox(position.path)) bad('The content of a text box cannot be edited');
  if (!block || typeof block.type !== 'string') bad('Only blocks can be inserted');
  let placed = block;
  if (placed.srcStart !== undefined || placed.srcEnd !== undefined) {
    // A block taken from elsewhere must not claim a place in the source of this story.
    placed = { ...placed };
    delete placed.srcStart;
    delete placed.srcEnd;
  }
  const container = position.path.slice(0, -1);
  let index = at.index;
  if (position.offset > 0 && position.offset < at.length) {
    const pieces = flatten(at.paragraph);
    const cut = caretIndex(pieces, position);
    draft.replace(position.story, container, index, 1, [
      paragraphWith(at.paragraph, rebuild(pieces.slice(0, cut)), withoutSection(at.paragraph.props)),
      paragraphWith(at.paragraph, rebuild(pieces.slice(cut)), at.paragraph.props, false),
    ]);
    index++;
  } else if (position.offset > 0) index++;
  const blocks = containerBlocks(doc, position.story, container);
  if (placed.wrap === undefined) {
    const wrap = wrapBetween(index > 0 ? blocks[index - 1].wrap : undefined, index < blocks.length ? blocks[index].wrap : undefined);
    if (wrap) placed = { ...placed, wrap };
  }
  const out: Block[] = [placed];
  let offset = 0;
  if (placed.type === 'table') {
    const empty = (): Paragraph => (placed.wrap ? { type: 'paragraph', props: EMPTY_PARA_PROPS, children: [], wrap: placed.wrap } : { type: 'paragraph', props: EMPTY_PARA_PROPS, children: [] });
    const before = neighbour(blocks, index, -1);
    const after = neighbour(blocks, index - 1, 1);
    if (before >= 0 && blocks[before].type === 'table') {
      out.unshift(empty());
      offset = 1;
    }
    if (after < 0 || blocks[after].type === 'table') out.push(empty());
  }
  draft.replace(position.story, container, index, 0, out);
  return index + offset;
}

/** Inserts a block (a paragraph, a table) at a position; see placeBlock. The caret goes to the start of the block, or of what follows it when it holds no paragraph. */
export function insertBlock(doc: DocModel, position: DocPosition, block: Block, options?: DocEditOptions): DocEdit {
  const build = (draft: Draft) => {
    const index = placeBlock(draft, position, block);
    const container = position.path.slice(0, -1);
    const path = paragraphIn(containerBlocks(doc, position.story, container), container, index, 1);
    return path ? { story: position.story, path, offset: 0 } : position;
  };
  // A block put in whole is not made a tracked insertion yet: refused where changes are tracked.
  return buildEdit(doc, build, options);
}
