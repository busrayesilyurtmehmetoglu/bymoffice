// Text commands: typing, deleting, breaking and joining paragraphs. Each
// returns the operations that make the change and the selection it leaves;
// nothing is changed until the caller applies them (see DocHistory).

import { paragraphText, EMPTY_PARA_PROPS, type Block, type CellProps, type DocModel, type ParaProps, type Paragraph, type RunItem, type Table, type TableCell, type TableRow } from '../../model/doc/types';
import { buildEdit, tableContextOf, type DocEdit, type DocEditOptions, type Draft } from './draft';
import { inList, patchRunProps, withoutNumbering, type RunFormatPatch } from './format';
import {
  boundary,
  cleanText,
  flatten,
  groupFor,
  inheritedProps,
  isAtom,
  itemPiece,
  offsetOf,
  paragraphWith,
  rebuild,
  removePieces,
  textPieces,
  wrapBetween,
  type Piece,
} from './pieces';
import {
  asRange,
  comparePositions,
  containerBlocks,
  inTextBox,
  nextCharOffset,
  nextWordOffset,
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

/** Paragraph properties without the section break: of the two halves of a paragraph only the second ends the section. */
function withoutSection(props: ParaProps): ParaProps {
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
function joinedProps(first: Paragraph, last: Paragraph): ParaProps {
  if (paragraphLength(first) === 0) return last.props;
  const section = last.props.section;
  const base = withoutSection(first.props);
  return section === undefined ? base : { ...base, section };
}

/** The first paragraph of a cell, emptied: what a cleared cell holds. */
function clearedBlocks(cell: TableCell): Block[] | undefined {
  let first: Paragraph | undefined;
  for (const block of cell.blocks) {
    if (block.type === 'paragraph') {
      first = block;
      break;
    }
  }
  if (first && cell.blocks.length === 1 && first.children.length === 0) return undefined;
  const paragraph: Paragraph = { type: 'paragraph', props: first ? withoutSection(first.props) : EMPTY_PARA_PROPS, children: [] };
  return [paragraph];
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

/**
 * Removes what lies between two positions (`start` before `end`) and returns
 * where the caret is afterwards.
 *
 * Within a paragraph the content goes. Across paragraphs the first one keeps
 * what is before the range and takes what the last one has after it; the
 * paragraphs and tables in between go. Where the range ends inside a table
 * it began outside of (or the other way round), and between cells of one
 * table, the cells in the range are emptied and the table stays as it is.
 * What is not content (bookmarks, comment marks, kept XML, fields the range
 * only partly covers) is never removed; of removed paragraphs it moves into
 * the paragraph that stays.
 */
export function deleteSpan(draft: Draft, start: DocPosition, end: DocPosition): DocPosition {
  const doc = draft.doc;
  const story = start.story;
  const a = resolvePosition(doc, start);
  const b = resolvePosition(doc, end);
  if (inTextBox(start.path) || inTextBox(end.path)) bad('The content of a text box cannot be edited');
  if (comparePositions(start, end) > 0) bad('The range is the wrong way round');

  let depth = 0;
  while (depth < start.path.length && depth < end.path.length && start.path[depth] === end.path[depth]) depth++;
  if (depth === start.path.length && depth === end.path.length) {
    const pieces = flatten(a.paragraph);
    const from = boundary(pieces, start.offset, 'start');
    const to = boundary(pieces, end.offset, 'caret');
    const offset = offsetOf(pieces, from);
    if (from < to) {
      const left = removePieces(pieces, from, to);
      if (left.length !== pieces.length) draft.setParagraph(story, start.path, paragraphWith(a.paragraph, rebuild(left)));
    }
    return { story, path: start.path.slice(), offset };
  }

  const level = depth - (depth % 3);
  const container = start.path.slice(0, level);
  if (depth % 3 !== 0) {
    // Two cells of one table: the cells of the rectangle between them are emptied.
    const blocks = containerBlocks(doc, story, container);
    const table = blocks[start.path[level]] as Table;
    const [r1, c1, r2, c2] = [start.path[level + 1], start.path[level + 2], end.path[level + 1], end.path[level + 2]];
    let cleared: Table;
    if (r1 === r2) cleared = tableCleared(table, (r, c) => r === r1 && c >= c1 && c <= c2);
    else {
      const first = cellSpan(table.rows[r1], c1);
      const last = cellSpan(table.rows[r2], c2);
      const left = Math.min(first[0], last[0]);
      const right = Math.max(first[1], last[1]);
      cleared = tableCleared(table, (r, c) => {
        if (r < r1 || r > r2) return false;
        const span = cellSpan(table.rows[r], c);
        return span[0] < right && span[1] > left;
      });
    }
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

  // What stays of the paragraphs in the range, in order, and the blocks without content between them.
  const pieces: Piece[] = [];
  let from = 0;
  let caretOffset = 0;
  if (startPara) {
    appendPieces(pieces, flatten(startPara));
    from = boundary(pieces, start.offset, 'start');
    caretOffset = offsetOf(pieces, from);
  }
  const kept: Block[] = [];
  let middle: Paragraph | undefined;
  for (let i = first + 1; i < last; i++) {
    const block = blocks[i];
    if (block.type === 'paragraph') {
      appendPieces(pieces, flatten(block));
      middle = block;
    } else if (block.type !== 'table') kept.push(block);
  }
  let to = pieces.length;
  if (endPara) {
    const tail = flatten(endPara);
    to += boundary(tail, end.offset, 'caret');
    appendPieces(pieces, tail);
  }
  const children = rebuild(removePieces(pieces, from, to));

  const out: Block[] = [];
  const startTable = startPara ? undefined : (blocks[first] as Table);
  const endTable = endPara ? undefined : (blocks[last] as Table);
  let caret: DocPosition = { story, path: start.path.slice(), offset: caretOffset };
  if (startTable) {
    const [r1, c1] = [start.path[level + 1], start.path[level + 2]];
    out.push(tableCleared(startTable, (r, c) => r > r1 || (r === r1 && c >= c1)));
  }
  if (startPara) {
    out.push(paragraphWith(startPara, children, endPara ? joinedProps(startPara, endPara) : startPara.props));
    appendBlocks(out, kept);
  } else if (endPara) {
    appendBlocks(out, kept);
    out.push(paragraphWith(endPara, children, endPara.props));
  } else {
    // Two tables that end up next to each other would become one: a paragraph that was between them stays, empty.
    if (middle) out.push(paragraphWith(middle, children, withoutSection(middle.props), false));
    appendBlocks(out, kept);
  }
  if (endTable) {
    const [r2, c2] = [end.path[level + 1], end.path[level + 2]];
    out.push(tableCleared(endTable, (r, c) => r < r2 || (r === r2 && c <= c2)));
  }
  draft.replace(story, container, first, last - first + 1, out);
  if (!startPara) {
    const cell = start.path.slice(0, level + 3);
    const path = paragraphIn(containerBlocks(doc, story, cell), cell, 0, 1);
    if (path) caret = { story, path, offset: 0 };
  }
  return caret;
}

/** Deletes the content of a range. The selection left is a caret where the range began. */
export function deleteRange(doc: DocModel, range: DocRange, _options?: DocEditOptions): DocEdit {
  return buildEdit(doc, (draft) => {
    const ordered = comparePositions(range.anchor, range.focus) <= 0;
    const start = ordered ? range.anchor : range.focus;
    const end = ordered ? range.focus : range.anchor;
    resolvePosition(doc, start);
    resolvePosition(doc, end);
    return samePosition(start, end) ? start : deleteSpan(draft, start, end);
  });
}

/** Most paragraphs one insertion makes: a paste beyond this is refused rather than left half done. */
const MAX_INSERTED_PARAGRAPHS = 1000000;

function insertAt(draft: Draft, position: DocPosition, text: string, format: RunFormatPatch | undefined): DocPosition {
  const doc = draft.doc;
  const at = resolvePosition(doc, position);
  if (inTextBox(position.path)) bad('The content of a text box cannot be edited');
  const paragraph = at.paragraph;
  const pieces = flatten(paragraph);
  const index = boundary(pieces, position.offset, 'caret');
  const offset = offsetOf(pieces, index);
  const inherited = inheritedProps(pieces, index, paragraph.props.markProps);
  const props = format ? patchRunProps(inherited, format) : inherited;
  const wrap = wrapBetween(index > 0 ? pieces[index - 1].wrap : undefined, index < pieces.length ? pieces[index].wrap : undefined);
  const group = groupFor(pieces, index, props, wrap);
  const lines = text.split('\n');
  if (lines.length > MAX_INSERTED_PARAGRAPHS) bad('The text has too many paragraphs');
  const story = position.story;
  const head = pieces.slice(0, index);
  appendPieces(head, textPieces(lines[0], props, wrap, group));
  if (lines.length === 1) {
    for (let i = index; i < pieces.length; i++) head.push(pieces[i]);
    draft.setParagraph(story, position.path, paragraphWith(paragraph, rebuild(head)));
    return { story, path: position.path.slice(), offset: offset + lines[0].length };
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
  return { story, path: container.concat([blockIndex + lines.length - 1]), offset: lastLine.length };
}

/**
 * Types or pastes text at a position, or over a range (whose content goes
 * first). "\n" ends a paragraph, a tab character becomes a tab and a vertical
 * tab ("\v", Shift+Enter) a line break. The text takes the formatting of the
 * text to its left, at the start of a paragraph of the text to its right, in
 * an empty paragraph of the paragraph mark; `format` is laid over that (the
 * "bold is on" of a caret that has no text yet).
 */
export function insertText(doc: DocModel, at: DocPosition | DocRange, text: string, format?: RunFormatPatch, _options?: DocEditOptions): DocEdit {
  if (typeof text !== 'string') bad('The text to insert must be a string');
  const range = asRange(at);
  const clean = cleanText(text);
  return buildEdit(doc, (draft) => {
    const ordered = comparePositions(range.anchor, range.focus) <= 0;
    const start = ordered ? range.anchor : range.focus;
    const end = ordered ? range.focus : range.anchor;
    resolvePosition(doc, start);
    resolvePosition(doc, end);
    const typing = samePosition(start, end);
    const caret = typing ? start : deleteSpan(draft, start, end);
    if (clean === '') return caret;
    const after = insertAt(draft, caret, clean, format);
    const selection = { anchor: after, focus: after };
    // One character typed at a caret: joins the characters typed before it into one undo step.
    if (!typing || nextCharOffset(clean, 0) !== clean.length || clean === '\n' || clean === '\t' || clean === '\v') return { selection };
    const space = clean === ' ' || clean === ' ';
    return { selection, mergeKey: 'type:' + start.story + ':' + start.path.join(','), seal: space };
  });
}

/** The nearest paragraph or table before (step -1) or after (1) a block of a list; blocks without content are passed over. */
function neighbour(blocks: Block[], index: number, step: 1 | -1): number {
  for (let i = index + step, tries = 0; i >= 0 && i < blocks.length && tries < 64; i += step, tries++) {
    if (blocks[i].type === 'paragraph' || blocks[i].type === 'table') return i;
  }
  return -1;
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

/** Removes the pieces [from, to) of a paragraph and returns the caret that leaves. */
function removeFrom(draft: Draft, position: DocPosition, paragraph: Paragraph, pieces: Piece[], from: number, to: number): DocPosition {
  const offset = offsetOf(pieces, from);
  const left = removePieces(pieces, from, to);
  if (left.length !== pieces.length) draft.setParagraph(position.story, position.path, paragraphWith(paragraph, rebuild(left)));
  return { story: position.story, path: position.path.slice(), offset };
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

/**
 * Backspace. Removes the character (or the word) before the caret; a drawing,
 * a note reference or a computed field there goes as one unit. At the start
 * of a paragraph, in Word's order: an item of a list leaves the list; else
 * an indent goes (first the first line indent, then the left one); else the
 * paragraph joins the one before it. Nothing happens at the start of a table
 * cell or of the story, and after a table.
 */
export function deleteBackward(doc: DocModel, position: DocPosition, unit: 'char' | 'word' = 'char', _options?: DocEditOptions): DocEdit {
  return buildEdit(doc, (draft) => {
    const at = resolvePosition(doc, position);
    if (inTextBox(position.path)) bad('The content of a text box cannot be edited');
    const paragraph = at.paragraph;
    const pieces = flatten(paragraph);
    const index = boundary(pieces, position.offset, 'caret');
    let k = index - 1;
    while (k >= 0 && pieces[k].len === 0 && !isAtom(pieces[k]) && pieces[k].lock === undefined) k--;
    if (k >= 0) {
      if (pieces[k].lock !== undefined) {
        const [from, to] = lockSpan(pieces, k);
        return removeFrom(draft, position, paragraph, pieces, from, to);
      }
      if (isAtom(pieces[k])) return removeFrom(draft, position, paragraph, pieces, k, k + 1);
      const text = paragraphText(paragraph);
      const end = offsetOf(pieces, k + 1);
      const begin = unit === 'word' ? previousWordOffset(text, end) : previousCharOffset(text, end);
      const from = boundary(pieces, begin, 'before');
      const to = boundary(pieces, end, 'after');
      return removeFrom(draft, position, paragraph, pieces, from, to);
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
    return joinParagraphs(draft, position.story, at.container, before, at.index);
  });
}

/**
 * Delete. Removes the character (or the word) after the caret, with the same
 * units as deleteBackward; at the end of a paragraph the next paragraph joins
 * this one. Nothing happens at the end of a table cell or of the story, and
 * before a table.
 */
export function deleteForward(doc: DocModel, position: DocPosition, unit: 'char' | 'word' = 'char', _options?: DocEditOptions): DocEdit {
  return buildEdit(doc, (draft) => {
    const at = resolvePosition(doc, position);
    if (inTextBox(position.path)) bad('The content of a text box cannot be edited');
    const paragraph = at.paragraph;
    const pieces = flatten(paragraph);
    const index = boundary(pieces, position.offset, 'caret');
    let k = index;
    while (k < pieces.length && pieces[k].len === 0 && !isAtom(pieces[k]) && pieces[k].lock === undefined) k++;
    if (k < pieces.length) {
      if (pieces[k].lock !== undefined) {
        const [from, to] = lockSpan(pieces, k);
        return removeFrom(draft, position, paragraph, pieces, from, to);
      }
      if (isAtom(pieces[k])) return removeFrom(draft, position, paragraph, pieces, k, k + 1);
      const text = paragraphText(paragraph);
      const begin = offsetOf(pieces, k);
      const end = unit === 'word' ? nextWordOffset(text, begin) : nextCharOffset(text, begin);
      const from = boundary(pieces, begin, 'before');
      const to = boundary(pieces, end, 'after');
      const caret = removeFrom(draft, position, paragraph, pieces, from, to);
      return { ...caret, offset: position.offset };
    }
    const after = neighbour(at.blocks, at.index, 1);
    if (after < 0 || at.blocks[after].type !== 'paragraph') return position;
    return joinParagraphs(draft, position.story, at.container, at.index, after);
  });
}

/**
 * Enter. The paragraph is cut in two at the position; both halves are
 * formatted like it, and a section break it carries stays with the second.
 * At the end of a paragraph whose style names a style for the paragraph that
 * follows (a heading), the new paragraph takes that style. In an empty list
 * item the paragraph leaves the list instead.
 *
 * UNSURE: whether Word keeps the direct formatting of the paragraph when it
 * switches to the following style; here the new paragraph has the style only.
 */
export function splitParagraph(doc: DocModel, position: DocPosition, _options?: DocEditOptions): DocEdit {
  return buildEdit(doc, (draft) => {
    const at = resolvePosition(doc, position);
    if (inTextBox(position.path)) bad('The content of a text box cannot be edited');
    const paragraph = at.paragraph;
    const where = tableContextOf(doc, position.story, position.path);
    if (at.length === 0 && inList(doc, paragraph.props, where)) {
      draft.setParagraph(position.story, position.path, paragraphWith(paragraph, paragraph.children, withoutNumbering(doc, paragraph.props, where)));
      return position;
    }
    const pieces = flatten(paragraph);
    const index = boundary(pieces, position.offset, 'caret');
    let second = paragraph.props;
    if (at.length > 0 && offsetOf(pieces, index) === at.length) {
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
    draft.replace(position.story, container, blockIndex, 1, [
      paragraphWith(paragraph, rebuild(pieces.slice(0, index)), withoutSection(paragraph.props)),
      paragraphWith(paragraph, rebuild(pieces.slice(index)), second, false),
    ]);
    return { story: position.story, path: container.concat([blockIndex + 1]), offset: 0 };
  });
}

const PAGE_BREAK_ITEM: RunItem = Object.freeze({ type: 'break' as const, kind: 'page' as const });

/** Ctrl+Enter. A page break goes in at the position and the paragraph ends after it: what follows starts the next page as a paragraph of its own. */
export function insertPageBreak(doc: DocModel, position: DocPosition, _options?: DocEditOptions): DocEdit {
  return buildEdit(doc, (draft) => {
    const at = resolvePosition(doc, position);
    if (inTextBox(position.path)) bad('The content of a text box cannot be edited');
    const paragraph = at.paragraph;
    const pieces = flatten(paragraph);
    const index = boundary(pieces, position.offset, 'caret');
    const props = inheritedProps(pieces, index, paragraph.props.markProps);
    const wrap = wrapBetween(index > 0 ? pieces[index - 1].wrap : undefined, index < pieces.length ? pieces[index].wrap : undefined);
    const head = pieces.slice(0, index);
    head.push(itemPiece(PAGE_BREAK_ITEM, props, wrap, groupFor(pieces, index, props, wrap)));
    const container = position.path.slice(0, -1);
    const blockIndex = position.path[position.path.length - 1];
    draft.replace(position.story, container, blockIndex, 1, [
      paragraphWith(paragraph, rebuild(head), withoutSection(paragraph.props)),
      paragraphWith(paragraph, rebuild(pieces.slice(index)), paragraph.props, false),
    ]);
    return { story: position.story, path: container.concat([blockIndex + 1]), offset: 0 };
  });
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
    const cut = boundary(pieces, position.offset, 'caret');
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
export function insertBlock(doc: DocModel, position: DocPosition, block: Block, _options?: DocEditOptions): DocEdit {
  return buildEdit(doc, (draft) => {
    const index = placeBlock(draft, position, block);
    const container = position.path.slice(0, -1);
    const path = paragraphIn(containerBlocks(doc, position.story, container), container, index, 1);
    return path ? { story: position.story, path, offset: 0 } : position;
  });
}
